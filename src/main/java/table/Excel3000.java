package table;

import com.google.common.collect.BiMap;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.HashBiMap;
import com.google.common.collect.Table;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public class Excel3000 {

  private Table<Integer, Integer, String> sheet = HashBasedTable.create();
  private static final Map<Character, Integer> rowMapping = new HashMap<>();
  private static final Pattern cellCoordPattern = Pattern.compile("^[A-Z]+[1-9]+$");
  private static final String ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  static {
    char[] alphabet = ALPHABET.toCharArray();
    for (int i = 0; i < alphabet.length; i++) {
      rowMapping.put(alphabet[i], i);
    }
  }

  private static int getRowNumber(String cell) throws IllegalArgumentException{
    if (!cellCoordPattern.matcher(cell).matches()){
      throw new IllegalArgumentException(cell + " could not be matched to a table index");
    }
    int row = 0;
    char[] rowChars = cell.chars()
        .filter(c -> Character.isAlphabetic(c) && Character.isUpperCase(c))
        .collect(StringBuilder::new, StringBuilder::appendCodePoint, StringBuilder::append)
        .toString()
        .toCharArray();
    for (int i = 0; i < rowChars.length; i++) {
      row += rowMapping.get(rowChars[i])+rowMapping.get(ALPHABET.charAt(ALPHABET.length()-1))*i;
    }
    return row;
  }

  public String getCellAt(int row, int col){
    return sheet.get(row,col);
  }
  static int[] getCoords(String cell){
    int[] coords = new int[2];
    int row = getRowNumber(cell);
    int col = Integer.parseInt(cell.substring(String.valueOf(row).length()));
    coords[0] = row;
    coords[1] = col;
    return coords;
  }
  public String getCellAt(String cell){
    int[] coords = getCoords(cell);
    return  getCellAt(coords[0], coords[1]);
  }

  public void setCell(int row, int col, String value){
    sheet.put(row, col, value);
  }

  public void setCell(String cell, String value) throws IllegalArgumentException{
    int[] coords = getCoords(cell);
    setCell(coords[0], coords[1],value);
  }


}

