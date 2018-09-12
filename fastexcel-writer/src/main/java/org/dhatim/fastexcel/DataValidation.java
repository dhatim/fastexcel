package org.dhatim.fastexcel;

import java.io.IOException;

/**
 *
 */
public interface DataValidation {

    /**
     * Write this dataValidation as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException;

    /**
     * Construct a new ListDataValidation
     *
     * @param range     The Range this validation is applied to
     * @param listRange The Range of the list this validation references
     */
    static ListDataValidation list(Range range, Range listRange) {
        return new ListDataValidation(range, listRange);
    }
}
