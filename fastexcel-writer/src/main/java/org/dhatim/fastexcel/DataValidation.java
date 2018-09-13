package org.dhatim.fastexcel;

import java.io.IOException;

interface DataValidation {

    /**
     * Write this dataValidation as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException;
}
