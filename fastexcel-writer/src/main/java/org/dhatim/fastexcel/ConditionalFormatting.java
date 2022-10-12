package org.dhatim.fastexcel;

import java.io.IOException;

public class ConditionalFormatting {
    private final Range range;
    private final ConditionalFormattingRule conditionalFormattingRule;

    /**
     * Constructor
     *
     * @param range                     The Range this validation is applied to
     * @param conditionalFormattingRule The Range of the list this validation references
     */
    ConditionalFormatting(Range range, ConditionalFormattingRule conditionalFormattingRule) {
        this.range = range;
        this.conditionalFormattingRule = conditionalFormattingRule;
    }
    
    /**
     * Get the conditional formatting rule of this conditional formatting.
     *
     * @return ConditionalFormattingRule conditional formatting rule.
     */
    ConditionalFormattingRule getConditionalFormattingRule() {
        return this.conditionalFormattingRule;
    }
    
    /**
     * Write this conditionalFormatting as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<conditionalFormatting sqref=\"").append(range.toString()).append("\">");
        conditionalFormattingRule.write(w);
        w.append("</conditionalFormatting>");
    }
}
