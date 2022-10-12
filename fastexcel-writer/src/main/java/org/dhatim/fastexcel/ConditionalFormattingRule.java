package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * A ConditionalFormattingRule defines a base class of conditional formatting rule for a worksheet
 */
public abstract class ConditionalFormattingRule {
    protected final String type;
    protected int priority;
    protected final boolean stopIfTrue;
    protected int dxfId;
    
    /**
     * Constructor
     *
     * @param type       Type of conditional formatting rule
     * @param stopIfTrue True indicates no rules with lower priority shall be applied over this rule, when this rule evaluates to true
     */
    ConditionalFormattingRule(String type, boolean stopIfTrue) {
        this.type = type;
        this.stopIfTrue = stopIfTrue;
    }

    /**
     * Set the priority
     *
     * @param priority The user interface display priority
     */
    void setPriority(int priority) {
        this.priority = priority;
    }

    /**
     * Set the dxfId
     *
     * @param dxfId The id of the style to apply when the conditional formatting rule criteria is met
     */
    void setDxfId(int dxfId) {
        this.dxfId = dxfId;
    }

    /**
     * Write this conditionalFormatting as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    abstract void write(Writer w) throws IOException;
}
