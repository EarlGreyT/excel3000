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
        expression = null;
      }
    }

  }

  private double evaluate(Set<CellValue> visitedCells) throws ArithmeticException {
    if(visitedCells.contains(this)){
      throw new IllegalStateException(this.expString);
    }
    visitedCells.add(this);
    for (String variable : variables) {
      CellValue neededExp = parent.getCellValueAt(variable);
      expression.setVariable(variable, neededExp.evaluate(visitedCells));
    }
    visitedCells.remove(this);
    return expression.evaluate();
  }

  public String showResult(Set<CellValue> visitedCells) {
    try {
      value = String.valueOf(evaluate(visitedCells));
      return value;
    }catch (IllegalStateException e){
      value="Cyclic Ref: " + e.getMessage();
      return value;
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
