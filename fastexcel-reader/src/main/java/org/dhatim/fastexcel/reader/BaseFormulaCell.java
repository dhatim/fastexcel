package org.dhatim.fastexcel.reader;

public class BaseFormulaCell {
    private CellAddress baseCelAddr;

    private String formula;
    private CellRangeAddress ref;


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
