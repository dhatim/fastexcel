package org.dhatim.fastexcel.reader;

public class ReadingOptions {
    public static final ReadingOptions DEFAULT_READING_OPTIONS = new ReadingOptions(false, false);
    private boolean withCellFormat;
    /**
     * If true, cell type is ERROR if it is not possible to parse cell value.
     * If false, an exception is throw when there is a parsing error
     */
    private boolean cellInErrorIfParseError;

    public ReadingOptions(boolean withCellFormat, boolean cellInErrorIfParseError) {
        this.withCellFormat = withCellFormat;
        this.cellInErrorIfParseError = cellInErrorIfParseError;
    }

    public boolean isWithCellFormat() {
        return withCellFormat;
    }

    public boolean isCellInErrorIfParseError() {
        return cellInErrorIfParseError;
    }
}
