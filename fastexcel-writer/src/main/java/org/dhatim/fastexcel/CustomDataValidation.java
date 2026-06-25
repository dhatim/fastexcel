package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * A CustomDataValidation defines a DataValidation for a worksheet of type = "custom"
 */
public class CustomDataValidation extends AbstractDataValidation<CustomDataValidation> {
    private final static String TYPE = "custom";
    private final Formula formula;

    /**
     * Constructor
     *
     * @param range     The Range this validation is applied to
     * @param formula   The Formula of the custom validation
     */
    CustomDataValidation(Range range, Formula formula) {
        super(range, TYPE);
        this.formula = formula;
    }

    /**
     * Write this dataValidation as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    protected void writeFormula(Writer w) throws IOException {
        w.append(formula.getExpression());
    }
}
