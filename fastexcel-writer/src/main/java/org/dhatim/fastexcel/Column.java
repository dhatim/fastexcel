package org.dhatim.fastexcel;

import java.util.Map;
import java.util.Objects;

/**
 * Definition of a column.
 */
class Column {

    /**
     * Worksheet where this column is defined.
     */
    private final Worksheet worksheet;
    /**
     * Position of the column
     */
    private final int colNumber;

    private int style;

    /**
     * Constructor
     * @param worksheet Worksheet where this column is defined.
     * @param colNumber Position of the column
     */
    Column(Worksheet worksheet, int colNumber) {
        this.worksheet = Objects.requireNonNull(worksheet);
        this.colNumber = Objects.requireNonNull(colNumber);
        this.style = 0;
    }

    static Column noStyle(Worksheet worksheet, int c) {
        return new Column(worksheet, c);
    }

    /**
     * Get parent worksheet.
     *
     * @return Parent worksheet.
     */
    public Worksheet getWorksheet() {
        return worksheet;
    }

    /**
     * Get column number.
     *
     * @return Column number.
     */
    public int getColNumber() {
        return colNumber;
    }

    /**
     * Get a new style setter for this column.
     *
     * @return Newly created style setter.
     */
    public ColumnStyleSetter style() {
        return new ColumnStyleSetter(this);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Column column = (Column) o;
        return colNumber == column.colNumber && Objects.equals(worksheet, column.worksheet);
    }

    @Override
    public int hashCode() {
        return Objects.hash(worksheet, colNumber);
    }

    /**
     * Return the style assigned to this column.
     *
     * @return style.
     */
    Integer getStyle() {
        return style;
    }

    /**
     * Apply new (merged) style to this column.
     *
     * @param stylesMap new styles map
     */
    void applyStyle(Map<Integer, Integer> stylesMap) {
        this.style = stylesMap.get(this.style);
    }
}
