/*
 * Copyright 2016 Dhatim.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.dhatim.fastexcel;

import java.io.Closeable;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.util.*;
import java.util.stream.IntStream;

/**
 * A worksheet is a set of cells.
 */
public class Worksheet implements Closeable {

    /**
     * Maximum number of rows in Excel.
     */
    public static final int MAX_ROWS = 1_048_576;

    /**
     * Maximum number of columns in Excel.
     */
    public static final int MAX_COLS = 16_384;

    /**
     * Maximum column width in Excel.
     */
    public static final int MAX_COL_WIDTH = 255;

     /**
     * Default column width in Excel.
     */
    public static final double DEFAULT_COL_WIDTH = 8.88671875;

    /**
     * Maximum row height in Excel.
     */
    public static final double MAX_ROW_HEIGHT = 409.5;

    private final Workbook workbook;
    private final String name;
    /**
     * List of rows. A row is an array of cells.
     * Flushed rows are null.
     */
    private final List<Cell[]> rows = new ArrayList<>();
    /**
     * Ranges of merged cells.
     */
    private final Set<Range> mergedRanges = new HashSet<>();
    /**
     * Matrix of merged cells.
     */
    private final DynamicBitMatrix mergedMatrix = new DynamicBitMatrix(MAX_COLS, MAX_ROWS);
    /**
     * List of conditional formattings for this worksheet
     */
    private final List<ConditionalFormatting> conditionalFormattings = new ArrayList<>();
    /**
     * List of DataValidations for this worksheet
     */
    private final List<DataValidation> dataValidations = new ArrayList<>();
    /**
     * List of ranges where shading to alternate rows is defined.
     */
    private final List<AlternateShading> alternateShadingRanges = new ArrayList<>();
    /**
     * List of ranges where shading to Nth rows is defined.
     */
    private final List<Shading> shadingRanges = new ArrayList<>();
    /**
     * List of rows to hide
     */
    private final Set<Integer> hiddenRows = new HashSet<>();

    /**
     * List of columns to hide
     */
    private final Set<Integer> hiddenColumns = new HashSet<>();
    /**
     * Array of column's group level
     */
    private final DynamicByteArray groupColums = new DynamicByteArray(MAX_COLS);
    /**
     * Array of rows's group level
     */
    private final DynamicByteArray groupRows = new DynamicByteArray(MAX_ROWS);

    /**
     * Map of columns and their widths
     */
    private final Map<Integer, Double> colWidths = new HashMap<>();

    /**
     * Map of columns and their representations with styles
     */
    private final Map<Integer, Column> colStyles = new HashMap<>();

    /**
     * Map of rows and their heights
     */
    private final Map<Integer, Double> rowHeights = new HashMap<>();

    final Comments comments = new Comments();

    final Map<String,Table> tables = new LinkedHashMap<>();

    private final DynamicBitMatrix tablesMatrix = new DynamicBitMatrix(MAX_COLS, MAX_ROWS);

    /**
     * Is this worksheet construction completed?
     */
    private boolean finished;

    /**
     * The visibility state of this sheet.
     */
    private VisibilityState visibilityState;

    /**
     * Whether grid lines are displayed
     */
    private boolean showGridLines = true;

    /**
     * Display the worksheet from right to left
     */
    private boolean rightToLeft = false;
    /**
     * Sheet view zoom percentage
     */
    private int zoomScale = 100;
    /**
     * Number of top rows that will remain frozen while scrolling.
     */
    private int freezeTopRows = 0;
    /**
     * Number of columns from the left that remain frozen while scrolling.
     */
    private int freezeLeftColumns = 0;
    /**
     * Page orientation [landscape / portrait] for the print preview setup.
     */
    private String pageOrientation = "portrait";
    /**
     * Paper size for the print preview setup.
     */
    private PaperSize paperSize = PaperSize.LETTER_PAPER;
    /**
     * Scaling factor for the print setup.
     */
    private int pageScale = 100;
    /**
     * Auto page breaks.
     */
    private Boolean autoPageBreaks = false;
    /**
     * Fit to page (true for fit to width/height).
     */
    private Boolean fitToPage = false;
    /**
     * Fit to width in the print setup.
     */
    private int fitToWidth = 1;
    /**
     * Fit to height in the print setup.
     */
    private int fitToHeight = 1;
    /**
     * First page number in the print setup.
     */
    private int firstPageNumber = 0;
    /**
     * Whether to use the firstPageNumber in the print setup.
     */
    private Boolean useFirstPageNumber=false;
    /**
     * Black and white mode in the print setup.
     */
    private Boolean blackAndWhite = false;
    /**
     * Header margin value in inches.
     */
    private float headerMargin = 0.3f;
    /**
     * Footer margin value in inches.
     */
    private float footerMargin = 0.3f;
    /**
     * Top margin value in inches.
     */
    private float topMargin = 0.75f;
    /**
     * Bottom margin value in inches.
     */
    private float bottomMargin = 0.75f;
    /**
     * Left margin value in inches.
     */
    private float leftMargin = 0.7f;
    /**
     * Right margin value in inches.
     */
    private float rightMargin = 0.7f;
    /**
     * Header map for left, central and right field text.
     */
    private final Map<Position, MarginalInformation> header = new LinkedHashMap<>();
    /**
     * Footer map for left, central and right field text.
     */
    private final Map<Position, MarginalInformation> footer = new LinkedHashMap<>();
    /**
     * Range of repeating rows for the print setup.
     * (Those rows will be repeated on each page when document is printed.)
     */
    private RepeatRowRange repeatingRows = null;
    /**
     * Range of repeating columns for the print setup.
     * (Those columns will be repeated on each page when document is printed.)
     */
    private RepeatColRange repeatingCols = null;
    /**
     * The hashed password that protects this sheet.
     */
    private String passwordHash;

    /**
     * Range of row where will be inserted auto filter
     */
    private Range autoFilterRange = null;

    private Relationships relationships = new Relationships(this);
    /**
     * List of named ranges.
     */
    private Map<String, Range> namedRanges = new LinkedHashMap<>();

    private Map<HyperLink, Ref> hyperlinkRanges = new LinkedHashMap<>();

    /**
     * The set of protection options that are applied on the sheet.
     */
    private Set<SheetProtectionOption> sheetProtectionOptions;

    private Writer writer;

    /**
     * Number of rows written to {@link #writer}.
     * Those rows are set to null in {@link #rows}
     */
    private int flushedRows = 0;

    private String tabColor;

    /**
     * Constructor.
     *
     * @param workbook Parent workbook.
     * @param name Worksheet name.
     */
    Worksheet(Workbook workbook, String name) {
        this.workbook = Objects.requireNonNull(workbook);
        this.name = Objects.requireNonNull(name);
    }

    /**
     * Get worksheet name.
     *
     * @return Worksheet name.
     */
    public String getName() {
        return name;
    }

    /**
     * Get repeating rows defined for the print setup.
     *
     * @return List representing a range of rows to be repeated
     *              on each page when printing.
     */
    public RepeatRowRange getRepeatingRows(){
        return repeatingRows;
    }

    /**
     * Get cell range that autofilter is applied to.
     *
     * @return Range of cells that autofilter is set to
     *             (null if autofilter is not set).
     */
    public Range getAutoFilterRange(){
        return autoFilterRange;
    }

    /**
     * Get repeating cols defined for the print setup.
     *
     * @return List representing a range of columns to be repeated
     *              on each page when printing.
     */
    public RepeatColRange getRepeatingCols(){
        return repeatingCols;
    }

    /**
     * Get a list of named ranges.
     *
     * @return Map containing named range entries
     *              where keys are the names and values are cell ranges.
     */
    public Map<String, Range> getNamedRanges() {
        return namedRanges;
    }

    /**
     * Get parent workbook.
     *
     * @return Parent workbook.
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Get the cell at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @return An existing or newly created cell.
     */
    Cell cell(int r, int c) {
        // Check limits
        if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) {
            throw new IllegalArgumentException();
        }
        flushedCheck(r);

        // Add null for missing rows.
        while (r >= rows.size()) {
            rows.add(null);
        }
        Cell[] row = rows.get(r);
        if (row == null) {
            int columns = Math.max(c + 1, (r > 0 && rows.get(r - 1) != null) ? rows.get(r - 1).length : (c + 1));
            row = new Cell[columns];
            rows.set(r, row);
        } else if (c >= row.length) {
            int columns = Math.max(c + 1, (r > 0 && rows.get(r - 1) != null) ? rows.get(r - 1).length : (c + 1));
            Cell[] tmp = new Cell[columns];
            System.arraycopy(row, 0, tmp, 0, row.length);
            row = tmp;
            rows.set(r, row);
        }
        if (row[c] == null) {
            row[c] = new Cell();
        }
        return row[c];
    }

    private void flushedCheck(int r) {
        if(r < flushedRows){
            throw new IllegalStateException("Row " + r + " already flushed from memory.");
        }
    }

    /**
     * Merge the cells within the given range.
     *
     * @param range Range of cells.
     */
    void merge(Range range) {
        if (!mergedMatrix.isConflict(range.getTop(),range.getLeft(),range.getBottom(),range.getRight())){
            if (mergedRanges.add(range)) {
                mergedMatrix.setRegion(range.getTop(),range.getLeft(),range.getBottom(),range.getRight());
            }
        }else {
            throw new IllegalArgumentException("Merge conflicted:" +range);
        }
    }

    /**
     * Apply shading to alternate rows in the given range.
     *
     * @param range Range of cells.
     * @param fill Shading fill pattern.
     */
    void shadeAlternateRows(Range range, Fill fill) {
        alternateShadingRanges.add(new AlternateShading(range, getWorkbook().cacheDifferentialFormat(new DifferentialFormat(null, null, fill, null, null, null))));
    }
    /**
     * Apply shading to Nth rows in the given range.
     *
     * @param range Range of cells.
     * @param fill Shading fill pattern.
     * @param eachNRows Shading row frequency.
     */
    void shadeRows(Range range, Fill fill, int eachNRows) {
        shadingRanges.add(new Shading(range, getWorkbook().cacheDifferentialFormat(new DifferentialFormat(null, null, fill, null, null, null)), eachNRows));
    }

    void addConditionalFormatting(ConditionalFormatting conditionalFormatting) {
        conditionalFormattings.add(conditionalFormatting);
    }

    void addValidation(DataValidation validation) {
        dataValidations.add(validation);
    }

    /**
     * Sets the visibility state of the sheet
     * <p>
     * This is done by setting the {@code state} attribute in the workbook.xml.
     * @param visibilityState New visibility state for this sheet.
     */
    public void setVisibilityState(VisibilityState visibilityState) {
        this.visibilityState = visibilityState;
    }

    public VisibilityState getVisibilityState() {
        return visibilityState;
    }

    /**
     * Hide the given row.
     *
     * @param row Zero-based row number
     */
    public void hideRow(int row) {
        hiddenRows.add(row);
    }

    /**
     * Show the given row.
     *
     * @param row Zero-based row number
     */
    public void showRow(int row) {
        hiddenRows.remove(row);
    }

    /**
     * Hide the given column.
     *
     * @param column Zero-based column number
     */
    public void hideColumn(int column) {
        hiddenColumns.add(column);
    }

    /**
     * Show the given column.
     *
     * @param column Zero-based column number
     */
    public void showColumn(int column) {
        hiddenColumns.remove(column);
    }

    /**
     * Keep this sheet in active tab.
     */
    public void keepInActiveTab() {
        int sheetIndex = workbook.getIndex(this);
        //tabs are indexed from 0, sheets are indexed from 1
        workbook.setActiveTab(sheetIndex - 1);
    }

    /**
     * Protects the sheet with a password. This method protects all the default {@link SheetProtectionOption}s and
     * 'sheet'. (Note that this is not very secure and only meant for discouraging changes.)
     * @param password The password to use.
     */
    public void protect(String password) {
        protect(password, SheetProtectionOption.DEFAULT_OPTIONS);
    }

    /**
     * Protects the sheet with a password. (Note that this is not very secure and only meant for discouraging changes.)
     * @param password The password to use.
     * @param options An array of all the {@link SheetProtectionOption}s to protect.
     */
    public void protect(String password, SheetProtectionOption... options) {
        final EnumSet<SheetProtectionOption> optionSet = EnumSet.noneOf(SheetProtectionOption.class);
        Collections.addAll(optionSet, options);
        protect(password, optionSet);
    }

    /**
     * Protects the sheet with a password. (Note that this is not very secure and only meant for discouraging changes.)
     * @param password The password to use.
     * @param options A {@link Set} of all the {@link SheetProtectionOption}s to protect.
     */
    public void protect(String password, Set<SheetProtectionOption> options) {
        if (password == null) {
            this.sheetProtectionOptions = null;
            this.passwordHash = null;
            return;
        }
        this.sheetProtectionOptions = options;
        this.passwordHash = hashPassword(password);
    }

    /**
     * Applies autofilter specifically to the given cell range
     * @param topRowNumber The first row (header) where filter will be initialized
     * @param leftCellNumber Left cell number where filter will be initialized
     * @param bottomRowNumber The last row (containing values) that will be included
     * @param rightCellNumber Right cell number where filter will be initialized
     */
    public void setAutoFilter(int topRowNumber, int leftCellNumber, int bottomRowNumber, int rightCellNumber) {
        autoFilterRange = new Range(this, topRowNumber, leftCellNumber, bottomRowNumber, rightCellNumber);
    }

    /**
     * Applies autofilter automatically based on provided header cells
     * @param rowNumber Row number
     * @param leftCellNumber Left cell number where filter will be initialized
     * @param rightCellNumber Right cell number where filter will be initialized
     */
    public void setAutoFilter(int rowNumber, int leftCellNumber, int rightCellNumber) {
        setAutoFilter(rowNumber, leftCellNumber, this.rows.size()-1, rightCellNumber);
    }

    /**
     * Removes auto filter from sheet. Does nothing if there wasn't any filter
     */
    public void removeAutoFilter() {
        autoFilterRange = null;
    }

    /**
     * Hash the password.
     * @param password The password to hash.
     * @return The password hash as a hex string (2 bytes)
     */
    private static String hashPassword(String password) {
        byte[] passwordCharacters = password.getBytes();
        int hash = 0;
        if (passwordCharacters.length > 0) {
            int charIndex = passwordCharacters.length;
            while (charIndex-- > 0) {
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters[charIndex];
            }
            // also hash with charcount
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= passwordCharacters.length;
            hash ^= (0x8000 | ('N' << 8) | 'K');
        }

        return Integer.toHexString(hash & 0xffff);
    }

    /**
     * Specify the width for the given column. Will autoSize by default.
     * <p>
     * The maximum column width in excel is 255. The colum width in
     * excel is the number of characters that can be displayed with the
     * standard font (first font in the workbook).
     * <p>
     * Note: The xml spec specifies additional padding for each cell
     * (Section 3.3.1.12 of the OOXML spec) which will result in slightly
     * less characters being displayed then what is given here.
     *
     * @param c     Zero-based column number
     * @param width The width of the column in character widths
     */
    public void width(int c, double width) {
        if (width > MAX_COL_WIDTH) {
            throw new IllegalArgumentException();
        }
        colWidths.put(c, width);
    }

    /**
     * Specify the custom row height for a row
     * <p> The maximum value for row height is <b>409.5</b> </p>
     * @param r Zero-based row number
     * @param height New row height
     */
    public void rowHeight(int r, double height) {
        if (height > MAX_ROW_HEIGHT) {
            throw new IllegalArgumentException();
        }
        rowHeights.put(r, height);
    }

    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value.
     */
    public void value(int r, int c, String value) {
        cell(r, c).setValue(workbook, value);
    }
    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value.
     */
    public void value(int r, int c, Number value) {
        cell(r, c).setValue(value);
    }
    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value.
     */
    public void value(int r, int c, Boolean value) {
        cell(r, c).setValue(value);
    }
    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value. Note Excel timestamps do not carry
     * any timezone information; {@link Date} values are converted to an Excel
     * serial number with the system timezone. If you need a specific timezone,
     * prefer passing a {@link ZonedDateTime}.
     */
    public void value(int r, int c, Date value) {
        cell(r, c).setValue(value);
    }
    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value. Note Excel timestamps do not carry
     * any timezone information; {@link Date} values are converted to an Excel
     * serial number with the system timezone. If you need a specific timezone,
     * prefer passing a {@link ZonedDateTime}.
     */
    public void value(int r, int c, LocalDateTime value) {
        cell(r, c).setValue(value);
    }
    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value. Note Excel timestamps do not carry
     * any timezone information; {@link Date} values are converted to an Excel
     * serial number with the system timezone. If you need a specific timezone,
     * prefer passing a {@link ZonedDateTime}.
     */
    public void value(int r, int c, LocalDate value) {
        cell(r, c).setValue(value);
    }
    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value.
     */
    public void value(int r, int c, ZonedDateTime value) {
        cell(r, c).setValue(value);
    }

    /**
     * Get the cell value (or formula) at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @return Cell value (or {@link Formula}).
     */
    public Object value(int r, int c) {
        flushedCheck(r);
        Cell[] row = r < rows.size() ? rows.get(r) : null;
        Cell cell = row == null || c >= row.length ? null : row[c];
        return cell == null ? null : cell.getValue();
    }

    public void hyperlink(int r, int c,HyperLink hyperLink) {
        value(r,c,hyperLink.getDisplayStr());
        this.addHyperlink(new Location(r,c),hyperLink);
    }

    /**
     * Set the cell formula at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param expression Cell formula expression.
     */
    public void formula(int r, int c, String expression) {
        cell(r, c).setFormula(expression);
    }

    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value.
     */
    public void inlineString(int r, int c, String value) {
        cell(r, c).setInlineString(value);
    }

    /**
     * Get a new style setter for a cell.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @return Newly created style setter.
     */
    public StyleSetter style(int r, int c) {
        return new Range(this, r, c, r, c).style();
    }

     /**
     * Get a new style setter for a column.
     *
     * @param c Zero-based column number.
     * @return Newly created style setter.
     */
    public ColumnStyleSetter style(int c) {
        Column column = colStyles.getOrDefault(c, new Column(this, c));
        colStyles.put(c, column);
        return column.style();
    }

    /**
     * Create a new range of cells. Note coordinates are reordered if necessary
     * to make sure {@code top} &lt;= {@code bottom} and {@code left} &lt;=
     * {@code right}.
     *
     * @param top Top row.
     * @param left Left column.
     * @param bottom Bottom row.
     * @param right Right column.
     * @return Newly created range.
     */
    public Range range(int top, int left, int bottom, int right) {
        return new Range(this, top, left, bottom, right);
    }

    /**
     * Write column definitions of this worksheet as an XML element.
     *
     * @param w Output writer.
     * @param maxCol Max number of columns.
     * @throws IOException If an I/O error occurs.
     */
    private void writeCols(Writer w, int maxCol) throws IOException {
        // Adjust column widths
        boolean started = false;
        for (int c = 0; c < maxCol; ++c) {
            double maxWidth = DEFAULT_COL_WIDTH;
            boolean bestFit = true;
            if (colWidths.containsKey(c)) {
                bestFit = false;
                maxWidth = colWidths.get(c);
            } else {
                for (int r = 0; r < rows.size(); ++r) {
                    boolean isCellInMergedRanges = mergedMatrix.get(r,c);
                    // Exclude merged cells from computation && hidden rows
                    Object o = hiddenRows.contains(r) || isCellInMergedRanges ? null : value(r, c);
                    if (o != null && !(o instanceof Formula)) {
                        int length = o.toString().length();
                        maxWidth = Math.max(maxWidth, (int) ((length * 7 + 10) / 7.0 * 256) / 256.0);
                    }
                }
            }
            boolean isHidden = hiddenColumns.contains(c);
            boolean hasStyle = colStyles.containsKey(c);
            boolean widthChanged = colWidths.containsKey(c) || maxWidth > DEFAULT_COL_WIDTH;
            int groupLevel = groupColums.get(c);
            if (widthChanged || isHidden || groupLevel != 0 || hasStyle) {
                if (!started) {
                    w.append("<cols>");
                    started = true;
                }
                Integer style = colStyles.getOrDefault(c, Column.noStyle(this, c)).getStyle();
                writeCol(w, c, maxWidth, bestFit, isHidden,groupLevel, style);
            }
        }
        if (started) {
            w.append("</cols>");
        }
    }


    /**
     * Write a column as an XML element.
     *
     * @param w Output writer.
     * @param columnIndex Zero-based column number.
     * @param maxWidth The maximum width
     * @param bestFit Whether or not this column should be optimized for fit
     * @param isHidden Whether or not this column is hidden
     * @param style Cached style index of the column
     * @throws IOException If an I/O error occurs.
     */
    private static void writeCol(Writer w, int columnIndex, double maxWidth, boolean bestFit, boolean isHidden, int groupLevel,
                                 int style) throws IOException {
        final int col = columnIndex + 1;
        w.append("<col min=\"").append(col).append("\" max=\"").append(col).append("\" width=\"")
                .append(Math.min(MAX_COL_WIDTH, maxWidth));
        w.append("\" outlineLevel=\"").append(groupLevel);
        w.append("\" customWidth=\"true\" bestFit=\"")
                .append(String.valueOf(bestFit));

        if (isHidden) {
            w.append("\" hidden=\"true");
        }
        w.append("\"");
        if (style > 0) {
            w.append(" style=\"").append(style).append("\"");
        }

        w.append("/>");
    }

    /**
     * Helper method to get a cell name from (x, y) cell position.
     * e.g. "B3" from cell position (2, 1)
     */
    private static String getCellMark(int row, int coll) {
        char columnLetter = (char) ('A' + coll);
        return String.valueOf(columnLetter) + String.valueOf(row+1);
    }

	@Override
	public void close() throws IOException {
		finish();
	}

    /**
     * Finish the construction of this worksheet. This creates the worksheet
     * file on the workbook's output stream. Rows and cells in this worksheet
     * are then destroyed.
     *
     * @throws IOException If an I/O error occurs.
     */
    public void finish() throws IOException {
        if (finished) {
            return;
        }
        flush();
        int index = workbook.getIndex(this);
        writer.append("</sheetData>");

        if (passwordHash != null) {
            writer.append("<sheetProtection password=\"").append(passwordHash).append("\" ");
            for (SheetProtectionOption option : SheetProtectionOption.values()) {
                if (option.getDefaultValue() != sheetProtectionOptions.contains(option)) {
                    writer.append(option.getName()).append("=\"").append(Boolean.toString(!option.getDefaultValue())).append("\" ");
                }
            }
            writer.append("/>");
        }
        if (autoFilterRange != null) {
            writer.append("<autoFilter ref=\"")
                    .append(autoFilterRange.toString())
                    .append("\">").append("</autoFilter>");
        }
        if (!mergedRanges.isEmpty()) {
            writer.append("<mergeCells>");
            for (Range r : mergedRanges) {
                writer.append("<mergeCell ref=\"").append(r.toString()).append("\"/>");
            }
            writer.append("</mergeCells>");
        }
        if (!conditionalFormattings.isEmpty()) {
            int priority = 1;
            for (ConditionalFormatting v: conditionalFormattings) {
                v.getConditionalFormattingRule().setPriority(priority++);
                v.write(writer);
            }
        }
        for (AlternateShading a : alternateShadingRanges) {
            a.write(writer);
        }
        for (Shading s : shadingRanges) {
            s.write(writer);
        }
        if (!dataValidations.isEmpty()) {
            writer.append("<dataValidations count=\"").append(dataValidations.size()).append("\">");
            for (DataValidation v: dataValidations) {
                v.write(writer);
            }
            writer.append("</dataValidations>");
        }
        if (!hyperlinkRanges.isEmpty()) {
            writer.append("<hyperlinks>");
            for (Map.Entry<HyperLink, Ref> hr : hyperlinkRanges.entrySet()) {
                HyperLink hyperLink = hr.getKey();
                writer.append("<hyperlink ");
                Ref ref = hr.getValue();
                writer.append("ref=\"" + ref.toString()+"\" ");
                if (hyperLink.getHyperLinkType().equals(HyperLinkType.EXTERNAL)) {
                    String rId = relationships.setHyperLinkRels(hyperLink.getLinkStr(), "External");
                    writer.append("r:id=\"" + rId +"\" ");
                }else{
                    writer.append("location=\"").append(hyperLink.getLinkStr()).append("\"");
                }
                writer.append("/>");
            }
            writer.append("</hyperlinks>");
        }
        /* set page margins for the print setup (see in print preview) */
        String margins = "<pageMargins bottom=\"" + bottomMargin +
                         "\" footer=\"" + footerMargin +
                         "\" header=\"" + headerMargin +
                         "\" left=\"" + leftMargin +
                         "\" right=\"" + rightMargin +
                         "\" top=\"" + topMargin + "\"/>";
        writer.append(margins);

        /* set page orientation for the print setup */
        writer.append("<pageSetup")
            .append(" paperSize=\"" + paperSize.xmlValue + "\"")
            .append(" scale=\"" + pageScale + "\"")
            .append(" fitToWidth=\"" + fitToWidth + "\"")
            .append(" fitToHeight=\"" + fitToHeight + "\"")
            .append(" firstPageNumber=\"" + firstPageNumber + "\"")
            .append(" useFirstPageNumber=\"" + useFirstPageNumber.toString() + "\"")
            .append(" blackAndWhite=\"" + blackAndWhite.toString() + "\"")
            .append(" orientation=\"" + pageOrientation + "\"")
            .append("/>");

        /* write to header and footer */
        writer.append("<headerFooter differentFirst=\"false\" differentOddEven=\"false\">");
        writer.append("<oddHeader>");
        for (MarginalInformation headerEntry : header.values()) {
            headerEntry.write(writer);
        }
        writer.append("</oddHeader>");
        writer.append("<oddFooter>");
        for (MarginalInformation footerEntry : footer.values()) {
            footerEntry.write(writer);
        }
        writer.append("</oddFooter></headerFooter>");


        if(!comments.isEmpty()) {
            writer.append("<drawing r:id=\"d\"/>");
            writer.append("<legacyDrawing r:id=\"v\"/>");
        }
        if (!tables.isEmpty()){
            writer.append("<tableParts count=\""+tables.size()+"\">");
            for (Map.Entry<String, Table> entry : tables.entrySet()) {
                writer.append("<tablePart r:id=\""+entry.getKey()+"\"/>");
            }
            writer.append("</tableParts>");
        }

        writer.append("</worksheet>");
        workbook.endFile();

        /* write comment files */
        if (!comments.isEmpty()) {
            workbook.writeFile("xl/comments" + index + ".xml", comments::writeComments);
            workbook.writeFile("xl/drawings/vmlDrawing" + index + ".vml", comments::writeVmlDrawing);
            workbook.writeFile("xl/drawings/drawing" + index + ".xml", comments::writeDrawing);
            relationships.setCommentsRels(index);
        }
        //write table files
        for (Map.Entry<String, Table> entry : tables.entrySet()) {
            Table table = entry.getValue();
            workbook.writeFile("xl/tables/table" + table.index + ".xml",table::write);
        }

        // write relationship files
        if (!relationships.isEmpty()) {
            workbook.writeFile("xl/worksheets/_rels/sheet"+index+".xml.rels",relationships::write);
        }
        // Free memory; we no longer need this data
        rows.clear();
        finished = true;
    }

    /**
     * Write all the rows currently in memory to the workbook's output stream.
     * Call this method periodically when working with huge data sets.
     * After calling {@link #flush()}, all the rows created so far become inaccessible.<br>
     * Notes:<br>
     * <ul>
     * <li>All columns must be defined before calling this method:
     * do not add or merge columns after calling {@link #flush()}.</li>
     * <li>When a {@link Worksheet} is flushed, no other worksheet can be flushed until {@link #close()} (or  the old fashion way {@link #finish()}) is called.</li>
     * </ul>
     *
     * @throws IOException If an I/O error occurs.
     */
    public void flush() throws IOException {
        if (writer == null) {
            int index = workbook.getIndex(this);
            writer = workbook.beginFile("xl/worksheets/sheet" + index + ".xml");
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            writer.append("<sheetPr filterMode=\"" + "false" + "\">" + (tabColor != null ? "<tabColor rgb=\""+tabColor+"\"/>" : "") + "<pageSetUpPr fitToPage=\"" + fitToPage + "\" autoPageBreaks=\"" + autoPageBreaks + "\"/></sheetPr>");
            writer.append("<dimension ref=\"A1\"/>");
            writer.append("<sheetViews><sheetView workbookViewId=\"0\"");
            if (!showGridLines) {
                writer.append(" showGridLines=\"false\"");
            }
            if (rightToLeft) {
                writer.append(" rightToLeft=\"true\"");
            }
            if (zoomScale != 100) {
                writer.append(" zoomScale=\"").append(zoomScale).append("\"");
            }
            writer.append(">");
            if (freezeLeftColumns > 0 || freezeTopRows > 0) {
                writeFreezePane(writer);
            }
            writer.append("</sheetView>");
            writer.append("</sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/>");
            final int nbCols = rows.stream().filter(Objects::nonNull).mapToInt(r -> r.length).max().orElse(0);
            final int maxHideCol = hiddenColumns.stream().mapToInt(a -> a).max().orElse(0);
            final int maxStyleCol = colStyles.values().stream().mapToInt(Column::getColNumber).max().orElse(0);
            final int maxNoZeroIndex = groupColums.getMaxNoZeroIndex();
            if (nbCols > 0 || !hiddenColumns.isEmpty()||maxNoZeroIndex!=-1 || !colStyles.isEmpty()) {
                int maxCol = Math.max(nbCols, Math.max(Math.max(maxHideCol,maxNoZeroIndex), maxStyleCol) + 1);
                writeCols(writer, maxCol);
            }
            writer.append("<sheetData>");
        }
        final int nbRows = rows.size();
        final int maxHideRow = hiddenRows.stream().mapToInt(a -> a).max().orElse(0);
        final int maxGroupRow = groupRows.getMaxNoZeroIndex();
        final int maxRow = Math.max(nbRows, Math.max(maxGroupRow,maxHideRow) + 1);
        for (int r = flushedRows; r < maxRow; ++r) {
            boolean notEmptyRow = r < rows.size();
            Cell[] row = notEmptyRow ? rows.get(r) : null;
            boolean isHidden = hiddenRows.contains(r);
            byte groupLevel = groupRows.get(r);
            if (row != null || isHidden || groupLevel != 0) {
                writeRow(writer, r, isHidden,groupLevel,
                        rowHeights.get(r), row);
            }
            if (notEmptyRow) {
                rows.set(r, null); // free flushed row data
            }
        }
        flushedRows = maxRow - 1;


        writer.flush();
    }

    /**
     * Writes corresponding pane definitions into XML and freezes pane.
     */
    private void writeFreezePane(Writer w) throws IOException {
        String activePane = freezeLeftColumns==0 ? "bottomLeft" : freezeTopRows==0 ? "topRight" : "bottomRight";
        String freezePane = "<pane xSplit=\"" + freezeLeftColumns +
                            "\" ySplit=\"" + freezeTopRows + "\" topLeftCell=\"" +
                            getCellMark(freezeTopRows, freezeLeftColumns) +
                            "\" activePane=\"" + activePane + "\" state=\"frozen\"/>";
        w.append(freezePane);
        String topLeftPane = "<selection pane=\"topLeft\" activeCell=\"" +
                             getCellMark(0, 0) +
                             "\" activeCellId=\"0\" sqref=\"" +
                             getCellMark(0, 0) + "\"/>";
        w.append(topLeftPane);
        if (freezeLeftColumns != 0) {
            String topRightPane = "<selection pane=\"topRight\" activeCell=\"" +
                                  getCellMark(0, freezeLeftColumns) +
                                  "\" activeCellId=\"0\" sqref=\"" +
                                  getCellMark(0, freezeLeftColumns) + "\"/>";
            w.append(topRightPane);
        }
        if (freezeTopRows !=0 ) {
            String bottomLeftPane = "<selection pane=\"bottomLeft\" activeCell=\"" +
                                    getCellMark(freezeTopRows, 0) +
                                    "\" activeCellId=\"0\" sqref=\"" +
                                    getCellMark(freezeTopRows, 0) + "\"/>";
            w.append(bottomLeftPane);
        }
        if (freezeLeftColumns != 0 && freezeTopRows != 0) {
            String bottomRightPane = "<selection pane=\"bottomRight\" activeCell=\"" +
                                     getCellMark(freezeTopRows, freezeLeftColumns) +
                                     "\" activeCellId=\"0\" sqref=\"" +
                                     getCellMark(freezeTopRows, freezeLeftColumns) + "\"/>";
            w.append(bottomRightPane);
        }
    }

    /**
     * Write a row as an XML element.
     *
     * @param w Output writer.
     * @param r Zero-based row number.
     * @param isHidden Whether or not this row is hidden
     * @param groupLevel Group level of row
     * @param rowHeight Row height value in points to be set if customHeight is true
     * @param row Cells in the row.
     * @throws IOException If an I/O error occurs.
     */
    private static void writeRow(Writer w, int r, boolean isHidden,byte groupLevel,
                                 Double rowHeight, Cell... row) throws IOException {
        w.append("<row r=\"").append(r + 1).append("\"");
        if (isHidden) {
            w.append(" hidden=\"true\"");
        }
        if(rowHeight != null) {
            w.append(" ht=\"")
             .append(rowHeight)
             .append("\"")
             .append(" customHeight=\"1\"");
        }
        if (groupLevel!=0){
            w.append(" outlineLevel=\"")
                    .append(groupLevel)
                    .append("\"");
        }
        w.append(">");
        if (null!=row) {
            for (int c = 0; c < row.length; ++c) {
                if (row[c] != null) {
                    row[c].write(w, r, c);
                }
            }
        }
        w.append("</row>");
    }

    /**
     * Assign a note/comment to a cell.
     * The comment popup will be twice the size of the cell and will be initially hidden.
     * <p>
     * Comments are stored in memory till call to {@link #close()} (or  the old fashion way {@link #finish()}) - calling {@link #flush()} does not write them to output stream.
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param comment Note text
     */
    public void comment(int r, int c, String comment) {
        comments.set(r, c, comment);
    }

    /**
     * Hide grid lines.
     */
    public void hideGridLines() {
        this.showGridLines = false;
    }

    /**
     * Display the worksheet from right to left
     */
    public void rightToLeft() {
        this.rightToLeft = true;
    }

    /**
     * Set sheet view zoom level in percent. Default is 100 (100%).
     * @param zoomPercent - zoom level from 10 to 400
     */
    public void setZoom(int zoomPercent) {
        if (10 <= zoomPercent && zoomPercent <= 400) {
            this.zoomScale = zoomPercent;
        }else{
            throw new IllegalArgumentException("zoom must be within 10 and 400 inclusive");
        }
    }

    public void setAutoPageBreaks(Boolean autoPageBreaks) {
        this.autoPageBreaks = autoPageBreaks;
    }

    public void setFitToPage(Boolean fitToPage) {
        this.fitToPage = fitToPage;
    }

    /**
     * Set freeze pane (rows and columns that remain when scrolling).
     * @param nLeftColumns - number of columns from the left that will remain frozen
     * @param nTopRows - number of rows from the top that will remain frozen
     */
    public void freezePane(int nLeftColumns, int nTopRows) {
        this.freezeLeftColumns = nLeftColumns;
        this.freezeTopRows = nTopRows;
    }

    /**
     * Unfreeze any frozen rows, or columns.
     */
    public void unfreeze() {
        this.freezeLeftColumns = 0;
        this.freezeTopRows = 0;
    }

    /**
     * Set header margin.
     * @param margin - header margin in inches
     */
    public void headerMargin(float margin) {
        this.headerMargin = margin;
    }

    /**
     * Set footer margin.
     * @param margin - footer page margin in inches
     */
    public void footerMargin(float margin) {
        this.footerMargin = margin;
    }

    /**
     * Set top margin.
     * @param margin - top page margin in inches
     */
    public void topMargin(float margin) {
        this.topMargin = margin;
    }

    /**
     * Set bottom margin.
     * @param margin - bottom page margin in inches
     */
    public void bottomMargin(float margin) {
        this.bottomMargin = margin;
    }

    /**
     * Set left margin.
     * @param margin - left page margin in inches
     */
    public void leftMargin(float margin) {
        this.leftMargin = margin;
    }

    /**
     * Set right margin.
     * @param margin - right page margin in inches
     */
    public void rightMargin(float margin) {
        this.rightMargin = margin;
    }

    /**
     * Set the page orientation.
     * @param orientation New page orientation for this worksheet
     */
    public void pageOrientation(String orientation) {
        this.pageOrientation = orientation;
    }
    /**
     * Set the paper size.
     * @param size New paper size for this worksheet
     */
    public void paperSize(PaperSize size) {
        this.paperSize = size;
    }
    /**
     * @param scale = scaling factor for the print setup (between 1 and 100)
     *
     */
    public void pageScale(int scale) {
        this.pageScale = scale;
    }
    /**
     * @param pageNumber - first page number (default: 0)
     */
    public void firstPageNumber(int pageNumber) {
        this.firstPageNumber = pageNumber;
        this.useFirstPageNumber = true;
    }

    public void fitToHeight(Short fitToHeight) {
        this.fitToPage = true;
        this.fitToHeight = fitToHeight;
    }

    public void fitToWidth(Short fitToWidth) {
        this.fitToPage = true;
        this.fitToWidth = fitToWidth;
    }

    public void printInBlackAndWhite() {
        this.blackAndWhite = true;
    }
    public void printInColor() {
        this.blackAndWhite = false;
    }

    public void repeatRows(int startRow, int endRow) {
        this.repeatingRows = new RepeatRowRange(startRow, endRow);
    }

    public void repeatRows(int row) {
        this.repeatingRows = new RepeatRowRange(row, row);
    }

    public void repeatCols(int startCol, int endCol) {
        this.repeatingCols = new RepeatColRange(startCol, endCol);
    }

    public void repeatCols(int col) {
        this.repeatingCols = new RepeatColRange(col, col);
    }

    /**
     * Set footer text.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     */
    public void footer(String text, Position position) {
        this.footer.put(position, new MarginalInformation(text, position));
    }

    /**
     * Set footer text with specified font size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontSize - integer describing font size
     */
    public void footer(String text, Position position, int fontSize) {
        this.footer.put(position, new MarginalInformation(text, position)
            .withFontSize(fontSize));
    }

    /**
     * Set footer text with specified font and size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontName - font name (e.g., "Arial")
     * @param fontSize - integer describing font size
     */
    public void footer(String text, Position position, String fontName, int fontSize) {
        this.footer.put(position, new MarginalInformation(text, position)
            .withFont(fontName)
            .withFontSize(fontSize));
    }

    /**
     * Set header text.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontName - font name (e.g., "Arial")
     * @param fontSize - integer describing font size
     */
    public void header(String text, Position position, String fontName, int fontSize) {
        this.header.put(position, new MarginalInformation(text, position)
            .withFont(fontName)
            .withFontSize(fontSize));
    }

    /**
     * Set header text with specified font size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontSize - integer describing font size
     */
    public void header(String text, Position position, int fontSize) {
        this.header.put(position, new MarginalInformation(text, position)
            .withFontSize(fontSize));
    }

    /**
     * Set header text with specified font and size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     */
    public void header(String text, Position position) {
        this.header.put(position, new MarginalInformation(text, position));
    }

    /**
     * Add the given range to this sheet's
     * list of named ranges under the provided name.
     * It will be visible when this sheet is open in the
     * cell range dropdown menu under the specified name.
     *
     * @param range Range of cells that needs to be named.
     * @param name String representing the given cell range's name.
     *
     */
    public void addNamedRange(Range range, String name) {
        this.namedRanges.put(name, range);
    }

    void addHyperlink(Ref ref, HyperLink hyperLink) {
        this.hyperlinkRanges.put(hyperLink, ref);
    }


    Table addTable(Range range, String... headers) {
        if (!tablesMatrix.isConflict(range.getTop(), range.getLeft(), range.getBottom(), range.getRight())) {
            int tableIndex = getWorkbook().nextTableIndex();
            String rId = relationships.setTableRels(tableIndex);
            Table table = new Table(tableIndex, range, headers);
            tables.put(rId, table);
            tablesMatrix.setRegion(range.getTop(), range.getLeft(), range.getBottom(), range.getRight());
            return table;
        } else {
            throw new IllegalArgumentException("Table conflicted:" + range);
        }
    }

    public void groupCols(int from , int to) {
        IntStream.rangeClosed(Math.min(from,to),Math.max(from,to)).forEach(groupColums::increase);
    }

    public void groupRows(int from , int to) {
        IntStream.rangeClosed(Math.min(from,to),Math.max(from,to)).forEach(groupRows::increase);
    }

    /**
     * @param rgbColor FFF381E0
     */
    public void setTabColor(String rgbColor) {
        this.tabColor = rgbColor;
    }
}
