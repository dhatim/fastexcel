package org.dhatim.fastexcel.reader;

public class ReadingOptions {
    public static final ReadingOptions DEFAULT_READING_OPTIONS = new ReadingOptions(false, false);
    private final boolean withCellFormat;
    private final boolean cellInErrorIfParseError;

    /**
     * @param withCellFormat          If true, extract cell formatting
     * @param cellInErrorIfParseError If true, cell type is ERROR if it is not possible to parse cell value.
     *                                If false, an exception is throw when there is a parsing error
     */
    public ReadingOptions(boolean withCellFormat, boolean cellInErrorIfParseError) {
        this.withCellFormat = withCellFormat;
        this.cellInErrorIfParseError = cellInErrorIfParseError;
    }

    /**
     * @return true for extract cell formatting
     */
    public boolean isWithCellFormat() {
        return withCellFormat;
    }

    /**
     * @return true for cell type is ERROR if it is not possible to parse cell value,
     * false for an exception is throw when there is a parsing error
     */
    public boolean isCellInErrorIfParseError() {
        return cellInErrorIfParseError;
    }
}
