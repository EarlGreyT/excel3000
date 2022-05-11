package table;

import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import net.objecthunter.exp4j.Expression;
import net.objecthunter.exp4j.ExpressionBuilder;

class CellValue {

  private Excel3000 parent;
  private String value;
  private String expString;
  private Expression expression;
  private Set<String> variables = new HashSet<>();

  private Pattern variablePattern;
  private Set<CellValue> callSet = new HashSet<>();
  private boolean visited = false;
  public CellValue(String value, Excel3000 parent) {
    this.parent = parent;
    this.value = value;
    this.expString = value;

    variablePattern = Pattern.compile("[A-Z]+[0-9]+");
    if(value == null){
      value="";
    }
    if (value.startsWith("=")) {
      extractVariables(value.substring(1));
      expression = new ExpressionBuilder(value.substring(1).replaceAll("\\$", "")).variables(
          variables).build();
    } else {
      if (!value.equals("")) {
        expression = new ExpressionBuilder(value).build();
      } else {
        expression = new ExpressionBuilder("0").build();
      }
    }

  }

  private static boolean checkCircle(CellValue callValue, CellValue caller, CellValue origin) {
    caller.visited = true;
    if (caller.callSet.contains(callValue)){
      return false;
    }
    if (origin.visited){
      return true;
    }
    for (CellValue value : callValue.callSet) {
      if (checkCircle(caller,callValue, origin)){
        return true;
      }
    }
    return false;
  }

  private double evaluate(CellValue caller) throws ArithmeticException {
    callSet.add(caller);
    if (checkCircle(this, caller, this)) {
      visited = false;
      caller.visited = false;
      throw new IllegalCallerException(callSet.stream().collect(StringBuilder::new,
          (StringBuilder sb, CellValue cv) -> sb.append(cv.expString + " "),
          StringBuilder::append).toString());
    }
    for (String variable : variables) {
      CellValue neededExp = parent.getCellValueAt(variable);
      expression.setVariable(variable, neededExp.evaluate(this));
    }
    visited = false;
    caller.visited = false;
    return expression.evaluate();
  }

  public String showResult() {
    try {
      value = String.valueOf(evaluate(this));
      return value;
    } catch (IllegalCallerException e) {
      return "#Cyclic Ref: " + e.getMessage();
    }

  }

  private void extractVariables(String value) {
    Matcher matcher = variablePattern.matcher(value);
    while (matcher.find()) {
      String group = matcher.group().replaceAll("\\$", "");
      variables.add(group);
    }

  }

  public String getValue() {
    return value;
  }

  public String getExpString() {
    return expString;
  }
}
