package table;


import com.google.common.collect.BiMap;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.HashBiMap;
import com.google.common.collect.Table;
import com.google.common.math.IntMath;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;


public class Excel3000 {

  private Table<Integer, Integer, String> sheet = HashBasedTable.create();
  private Table<Integer, Integer, CellValue> sheetValues = HashBasedTable.create();
  private static final BiMap<Character, Integer> rowMapping = HashBiMap.create();
  private static final Pattern cellCoordPattern = Pattern.compile("^[A-Z]+[1-9]+$");
  private static final char refSign ='$';
  private static final String ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  static {
    char[] alphabet = ALPHABET.toCharArray();
    for (int i = 0; i < alphabet.length; i++) {
      rowMapping.put(alphabet[i], i);
    }
  }

  private static int getRowNumber(String cell) throws IllegalArgumentException {
    if (!cellCoordPattern.matcher(cell).matches()) {
      throw new IllegalArgumentException(cell + " could not be matched to a table index");
    }
    int row = 0;
    char[] rowChars = cell.chars()
        .filter(c -> Character.isAlphabetic(c) && Character.isUpperCase(c))
        .collect(StringBuilder::new, StringBuilder::appendCodePoint, StringBuilder::append)
        .toString()
        .toCharArray();
    for (int i = 0; i < rowChars.length; i++) {
      // Hexadecimal 111 = 1*16^2+1*16^1+1*16^0
      // but our base is the length of the ALPHABET.
      row += rowMapping.get(rowChars[i]) * IntMath.pow(
          1 + rowMapping.get(ALPHABET.charAt(ALPHABET.length() - 1)), rowChars.length-(i+1));
    }
    return row;
  }

  public String getCellAt(int row, int col) {
    return sheet.get(row, col);
  }

  static int[] getCoords(String cell) {
    int[] coords = new int[2];
    int row = getRowNumber(cell);
    int col = (int) cell.chars().filter(Character::isDigit).count();
    coords[0] = row;
    coords[1] = col;
    return coords;
  }

  public String getCellAt(String cell) {
    int[] coords = getCoords(cell);
    return getCellAt(coords[0], coords[1]);
  }

  public void setCell(int row, int col, String value) {
    sheet.put(row, col, value);
  }

  public void setCell(String cell, String value) throws IllegalArgumentException {
    int[] coords = getCoords(cell);
    setCell(coords[0], coords[1], value);
  }


  public CellValue getCellValueAt(String variable) {
    int[] coords = getCoords(variable);
    return getCellValueAt(coords[0], coords[1]);
  }
  public void evaluate(){
    for (int i =0; i<sheet.rowMap().size();i++){
      for (int j=0; j<sheet.columnMap().size();j++){
        sheet.put(i,j, getCellValueAt(i,j).showResult());
      }
    }
  }

  @Override
  public String toString() {
    return sheet.toString();
  }

  private CellValue getCellValueAt(int row, int col) {
    CellValue value = sheetValues.contains(row,col) ?
        sheetValues.get(row,col)
        : new CellValue(getCellAt(row,col),refSign,this);
    sheetValues.put(row,col, value);
    return value;
  }
}

