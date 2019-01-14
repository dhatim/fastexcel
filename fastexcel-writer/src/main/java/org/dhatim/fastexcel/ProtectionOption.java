package org.dhatim.fastexcel;

/**
 * Represents an attribute on the &lt;protection&gt; xml-tag.
 */
public enum ProtectionOption {

    /**
     * A boolean value indicating if the cell is hidden. When the cell is hidden and the sheet on which the cell resides
     * is protected, then the cell value will be displayed in the cell grid location, but the contents of the cell will
     * not be displayed in the formula bar. This is true for all types of cell content, including formula, text, or
     * numbers.
     *
     * Therefore the cell A4 may contain a formula "=SUM(A1:A3)", but if the cell protection property of A4 is marked as
     * hidden, and the sheet is protected, then the cell should display the calculated result (for example, "6"), but
     * will not display the formula used to calculate the result.
     */
    HIDDEN("hidden"),

    /**
     * A boolean value indicating if the cell is locked. When cells are marked as "locked" and the sheet is protected,
     * then the options specified in the Sheet Part's &lt;sheetProtection&gt; element (§3.3.1.81) are prohibited for
     * these cells.
     */
    LOCKED("locked");

    private final String name;

    /**
     * Constructor that sets the name.
     *
     * @param name The name of the protection option.
     */
    ProtectionOption(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }
}
