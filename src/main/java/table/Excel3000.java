package table;


import com.google.common.collect.BiMap;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.HashBiMap;
import com.google.common.collect.Table;
import com.google.common.math.IntMath;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Pattern;


public class Excel3000 {
  private static final Set<CellValue> visitedCells = new HashSet<>();
  private final Table<Integer, Integer, String> sheet = HashBasedTable.create();
  private final Table<Integer, Integer, CellValue> sheetValues = HashBasedTable.create();
  private static final BiMap<Character, Integer> rowMapping = HashBiMap.create();
  public static final String CELL_CORD_REGEX = "^[A-Z]+[1-9][0-9]*$";
  private static final Pattern cellCoordPattern = Pattern.compile(CELL_CORD_REGEX);

  private static final String ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  static {
    char[] alphabet = ALPHABET.toCharArray();
    for (int i = 0; i < alphabet.length; i++) {
      rowMapping.put(alphabet[i], i);
    }
  }

  private static String getRowName(int row) {
    ArrayList<Integer> remainders = new ArrayList<>();
    int dividend = row;
    while (dividend >= 0) {
      remainders.add(dividend % ALPHABET.length());
      dividend = IntMath.divide(dividend, ALPHABET.length(), RoundingMode.DOWN);
    }
    Collections.reverse(remainders);
    return remainders.stream().collect(StringBuilder::new,
        ((StringBuilder sb, Integer i) -> sb.append(rowMapping.inverse().get(i))),
        StringBuilder::append).toString();
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
          rowMapping.get(ALPHABET.charAt(ALPHABET.length() - 1)), rowChars.length - (i + 1));
    }
    return row;
  }

  public String getCellAt(int row, int col) {
    String value = sheet.get(row, col);
    if (value == null) {
      value = "";
    }
    return value;
  }

  static int[] getCoords(String cell) {
    int[] coords = new int[2];
    int row = getRowNumber(cell);
    int col = Integer.parseInt(cell.chars().filter(Character::isDigit)
        .collect(StringBuilder::new, StringBuilder::appendCodePoint, StringBuilder::append)
        .toString());
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
    sheetValues.put(row, col, new CellValue(value, this));
  }

  public void setCell(String cell, String value) throws IllegalArgumentException {
    int[] coords = getCoords(cell);
    setCell(coords[0], coords[1], value);
  }


  public CellValue getCellValueAt(String variable) {
    int[] coords = getCoords(variable);
    return getCellValueAt(coords[0], coords[1]);
  }
  public void devaluate() {
    for (Integer i : sheet.rowKeySet()) {
      for (Integer j : sheet.columnKeySet()) {
        if (sheet.get(i, j) != null) {
          CellValue cellValue = getCellValueAt(i, j);
          if (cellValue != null) {
            sheet.put(i, j, cellValue.getValue());
          }
        }
      }
    }
  }
  public void evaluate() {
    for (Integer i : sheet.rowKeySet()) {
      for (Integer j : sheet.columnKeySet()) {
        if (sheet.get(i, j) != null) {
          CellValue cellValue = getCellValueAt(i, j);
          if (cellValue != null) {
            sheet.put(i, j, cellValue.showResult(visitedCells));
          }
        }
      }
    }
  }

  public Table<Integer, Integer, String> getSheet() {
    return sheet;
  }

  @Override
  public String toString() {
    return sheet.toString();
  }

  private CellValue getCellValueAt(int row, int col) {
    return sheetValues.get(row, col);
  }
}

