package writer;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import table.Excel3000;

public class XlsxWriter {

  public static void writeToExcellFile(String name, Excel3000 table) throws IOException {
    try (Workbook wb = new XSSFWorkbook(); FileOutputStream fOut = new FileOutputStream(name)) {
      Sheet s = wb.createSheet();
      for (Integer col : table.getSheet().columnKeySet()) {
        Row r = s.createRow(col - 1);
        for (Integer row : table.getSheet().rowKeySet()) {
          Cell c = r.createCell(row);
          String cellValue = table.getCellAt(row, col);
          if (cellValue.chars().allMatch(Character::isDigit) && !cellValue.equals("")) {
            c.setCellValue(Double.parseDouble(cellValue));
          } else {
            if (cellValue.startsWith("=")) {
              c.setCellFormula(cellValue.substring(1));
            } else {
              c.setCellValue(cellValue);
            }
          }
        }
      }
      wb.write(fOut);
    }

  }
}
