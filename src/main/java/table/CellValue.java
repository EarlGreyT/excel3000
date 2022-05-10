package table;

import java.util.HashSet;
import java.util.Set;
import java.util.regex.Pattern;
import net.objecthunter.exp4j.Expression;
import net.objecthunter.exp4j.ExpressionBuilder;

public class CellValue {
  private Excel3000 parent;
  private String value;
  private Expression expression;
  private Set<String> variables = new HashSet<>();
  private char refSign;
  private Pattern variablePattern = Pattern.compile("^" + refSign + "[A-Z]+[1-9][0-9]*$");

  public CellValue(String value, char refSign, Excel3000 parent) {
    this.parent = parent;
    this.value = value;
    this.refSign = refSign;
    if (value.startsWith("=")) {
      extractVariables(value.substring(1));
      expression = new ExpressionBuilder(value.substring(1).replaceAll("$","")).variables(variables).build();
    } else {
      expression = new ExpressionBuilder(value).build();
    }
  }

  private double evaluate(){
    for (String variable : variables) {
      expression.setVariable(variable, parent.getCellValueAt(variable).evaluate());
    }
    return expression.evaluate();
  }
  public String showResult(){
    double result = evaluate();
    expression = new ExpressionBuilder(String.valueOf(result)).build();
    value = String.valueOf(result);
    return value;
  }
  private void extractVariables(String substring) {
    variablePattern.matcher(substring).results().forEach(
        matchResult -> variables.add(matchResult.group().replace(String.valueOf(refSign), "")));
  }

}
