package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * A ListDataValidation defines a DataValidation for a worksheet of type = "list"
 */
public class ListDataValidation extends AbstractDataValidation<ListDataValidation> {
    private final static String TYPE = "list";
    private final Range listRange;

    /**
     * Constructor
     *
     * @param range     The Range this validation is applied to
     * @param listRange The Range of the list this validation references
     */
    ListDataValidation(Range range, Range listRange) {
        super(range, TYPE);
        this.listRange = listRange;
    }

    /**
     * Write this dataValidation as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    protected void writeFormula(Writer w) throws IOException {
        w.append("'").appendEscaped(listRange.getWorksheet().getName())
                .append("'!")
                .append(listRange.toAbsoluteString());
    }
}
