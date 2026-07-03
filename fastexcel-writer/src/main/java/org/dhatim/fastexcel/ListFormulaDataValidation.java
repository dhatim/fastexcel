package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * A ListDataValidation defines a DataValidation for a worksheet of type = "list"
 */
public class ListFormulaDataValidation extends AbstractDataValidation<ListFormulaDataValidation> {
    private final static String TYPE = "list";
    private final Formula formula;

    /**
     * Constructor
     *
     * @param range     The Range this validation is applied to
     * @param formula   The Formula of this validation to retrieve the list
     */
    ListFormulaDataValidation(Range range, Formula formula) {
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
