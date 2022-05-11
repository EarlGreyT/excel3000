package table;

public class Main {

  public static void main(String[] args) {
    Excel3000 excel3000 = new Excel3000();
    //circle

    excel3000.setCell("A0","4");

    excel3000.setCell("A1", "=$A0+4");
    excel3000.setCell("A2", "=$A6+4");
    excel3000.setCell("A4", "=$A2+4");
    excel3000.setCell("A6", "=$B2+4");
    excel3000.setCell("B2" ,"=$A4+1");
    excel3000.evaluate();
    System.out.println(excel3000);

    //no circle
    excel3000.setCell("A0","4");
    excel3000.setCell("A1", "=$A0+4");
    excel3000.setCell("A2", "=$A6+4");
    excel3000.setCell("A4", "=$A1+4");
    excel3000.setCell("A5", "=$A2+4");
    excel3000.setCell("A6", "=$A4+4");
    excel3000.evaluate();
    System.out.println(excel3000);
  }
}
