package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * A ConditionalFormattingExpressionRule defines a conditional formatting rule for a worksheet of type = "expression"
 */
public class ConditionalFormattingExpressionRule extends ConditionalFormattingRule {
    protected final static String TYPE = "expression";
    protected final String expression;

    /**
     * Constructor
     *
     * @param expression When the expression evaluates to true, the specified style is applied
     * @param stopIfTrue True indicates no rules with lower priority shall be applied over this rule, when this rule evaluates to true
     */
    public ConditionalFormattingExpressionRule(String expression, boolean stopIfTrue) {
        super(TYPE, stopIfTrue);
        this.expression = expression;
    }

    /**
     * Write this conditionalFormatting as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    @Override
    public void write(Writer w) throws IOException {
        w
            .append("<cfRule type=\"").append(TYPE).append("\" priority=\"").append(priority).append("\" stopIfTrue=\"").append(stopIfTrue ? "1" : "0").append("\" dxfId=\"").append(dxfId).append("\">")
            .append("<formula>").append(expression).append("</formula>")
            .append("</cfRule>");
    }
}
