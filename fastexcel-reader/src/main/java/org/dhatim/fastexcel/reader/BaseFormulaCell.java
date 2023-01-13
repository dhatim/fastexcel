package org.dhatim.fastexcel.reader;

public class BaseFormulaCell {
    private final CellAddress baseCelAddr;

    private final String formula;
    private final CellRangeAddress ref;


    public BaseFormulaCell(CellAddress baseCelAddr, String formula, CellRangeAddress ref) {
        this.baseCelAddr = baseCelAddr;
        this.formula = formula;
        this.ref = ref;
    }

    public CellAddress getBaseCelAddr() {
        return baseCelAddr;
    }

    public String getFormula() {
        return formula;
    }

    public CellRangeAddress getRef() {
        return ref;
    }
}
