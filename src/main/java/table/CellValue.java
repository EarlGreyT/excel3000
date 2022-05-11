package table;

import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import net.objecthunter.exp4j.Expression;
import net.objecthunter.exp4j.ExpressionBuilder;

class CellValue {

  public static final String VARIABLE_PATTERN = "[A-Z]+[1-9][0-9]*";
  private final Excel3000 parent;
  private String value;
  private final String expString;
  private final Expression expression;
  private final Set<String> variables = new HashSet<>();

  private final Pattern variablePattern;

  public CellValue(String value, Excel3000 parent) {
    this.parent = parent;
    this.value = value;
    this.expString = value;
    variablePattern = Pattern.compile(VARIABLE_PATTERN);
    if (value == null) {
      value = "";
    }
    if (value.startsWith("=")) {
      extractVariables(value.substring(1));
      expression = new ExpressionBuilder(value.substring(1)).variables(
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
    if (visitedCells.contains(this)) {
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
    } catch (IllegalStateException e) {
      value = "Cyclic Ref: " + e.getMessage();
      return value;
    }

  }

  private void extractVariables(String value) {
    Matcher matcher = variablePattern.matcher(value);
    while (matcher.find()) {
      String group = matcher.group();
      variables.add(group);
    }
  }

}
