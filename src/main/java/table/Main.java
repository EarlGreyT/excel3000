package table;

import java.io.IOException;
import writer.XlsxWriter;

public class Main {

  public static void main(String[] args) throws IOException {
    Excel3000 excel3000 = new Excel3000();
    Excel3000 excel3001 = new Excel3000();
    //circle

    excel3000.setCell("A1","4");

    excel3000.setCell("A2", "=A1+4");
    excel3000.setCell("A3", "=A7+4");
    excel3000.setCell("A5", "=A3+4");
    excel3000.setCell("A6", "=A2+4");
    excel3000.setCell("A7", "=B3+4");
    excel3000.setCell("B3" ,"=A5+1");
    XlsxWriter.writeToExcellFile("circleNoEval.xlsx",excel3000);
    excel3000.evaluate();
    XlsxWriter.writeToExcellFile("circleEval.xlsx",excel3000);
    System.out.println(excel3000);

    //no circle
    excel3001.setCell("A1","4");
    excel3001.setCell("A2", "=A3+4");
    excel3001.setCell("A3", "=A6+4");
    excel3001.setCell("A4", "=A1+4");
    excel3001.setCell("A5", "=A2+4");
    excel3001.setCell("A6", "=A4+4");
    XlsxWriter.writeToExcellFile("NoCircleNoEval.xlsx",excel3001);
    excel3001.evaluate();
    XlsxWriter.writeToExcellFile("NCircleEval.xlsx",excel3001);
    excel3001.devaluate();
    XlsxWriter.writeToExcellFile("NCircleDEval.xlsx",excel3001);
    System.out.println(excel3001);
  }
}
