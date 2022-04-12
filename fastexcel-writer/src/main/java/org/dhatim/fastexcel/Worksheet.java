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

import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.util.*;

/**
 * A worksheet is a set of cells.
 */
public class Worksheet {

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
     * Map of columns and their widths
     */
    private final Map<Integer, Double> colWidths = new HashMap<>();

    /**
     * Map of rows and their heights
     */
    private final Map<Integer, Double> rowHeights = new HashMap<>();

    final Comments comments = new Comments();

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
    private Map<Position,String> header = new LinkedHashMap();
    /**
     * Footer map for left, central and right field text.
     */
    private Map<Position,String> footer = new LinkedHashMap();
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
    /**
     * List of named ranges.
     */
    private Map<String, Range> namedRanges = new LinkedHashMap();

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
        mergedRanges.add(range);
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
     * Check if the given cell is within merged ranges.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @return {@code true} if the cell is within merged ranges, {@code false}
     * otherwise.
     */
    private boolean isCellInMergedRanges(int r, int c) {
        return mergedRanges.stream().anyMatch(range -> range.contains(r, c));
    }

    /**
     * Write column definitions of this worksheet as an XML element.
     *
     * @param w Output writer.
     * @param nbCols Number of columns.
     * @throws IOException If an I/O error occurs.
     */
    private void writeCols(Writer w, int nbCols) throws IOException {
        // Adjust column widths
        boolean started = false;
        for (int c = 0; c < nbCols; ++c) {
            double maxWidth = 0;
            boolean bestFit = true;
            if (colWidths.containsKey(c)) {
                bestFit = false;
                maxWidth = colWidths.get(c);
            } else {
                for (int r = 0; r < rows.size(); ++r) {
                    // Exclude merged cells from computation && hidden rows
                    Object o = hiddenRows.contains(r) || isCellInMergedRanges(r, c) ? null : value(r, c);
                    if (o != null && !(o instanceof Formula)) {
                        int length = o.toString().length();
                        maxWidth = Math.max(maxWidth, (int) ((length * 7 + 10) / 7.0 * 256) / 256.0);
                    }
                }
            }
            if (maxWidth > 0) {
                if (!started) {
                    w.append("<cols>");
                    started = true;
                }
                writeCol(w, c, maxWidth, bestFit, hiddenColumns.contains(c));
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
     * @param isHidden Whether or not this row is hidden
     * @throws IOException If an I/O error occurs.
     */
    private static void writeCol(Writer w, int columnIndex, double maxWidth, boolean bestFit, boolean isHidden) throws IOException {
        final int col = columnIndex + 1;
        w.append("<col min=\"").append(col).append("\" max=\"").append(col).append("\" width=\"")
                .append(Math.min(MAX_COL_WIDTH, maxWidth)).append("\" customWidth=\"true\" bestFit=\"")
                .append(String.valueOf(bestFit));

        if (isHidden) {
            w.append("\" hidden=\"true");
        }

        w.append("\"/>");
    }

    /**
     * Helper method to get a cell name from (x, y) cell position.
     * e.g. "B3" from cell position (2, 1)
     */
    private static String getCellMark(int row, int coll) {
        char columnLetter = (char) ('A' + coll);
        return String.valueOf(columnLetter) + String.valueOf(row+1);
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
        writer.append("</sheetData>");

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
        if (!dataValidations.isEmpty()) {
            writer.append("<dataValidations count=\"").append(dataValidations.size()).append("\">");
            for (DataValidation v: dataValidations) {
                v.write(writer);
            }
            writer.append("</dataValidations>");
        }
        for (AlternateShading a : alternateShadingRanges) {
            a.write(writer);
        }
        for (Shading s : shadingRanges) {
            s.write(writer);
        }

        if (passwordHash != null) {
            writer.append("<sheetProtection password=\"").append(passwordHash).append("\" ");
            for (SheetProtectionOption option : SheetProtectionOption.values()) {
                if (option.getDefaultValue() != sheetProtectionOptions.contains(option)) {
                    writer.append(option.getName()).append("=\"").append(Boolean.toString(!option.getDefaultValue())).append("\" ");
                }
            }
            writer.append("/>");
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
            .append(" paperSize=\"1\"")
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
        if (header.get(Position.LEFT) != null) writer.append(header.get(Position.LEFT));
        if (header.get(Position.CENTER) != null) writer.append(header.get(Position.CENTER));
        if (header.get(Position.RIGHT) != null) writer.append(header.get(Position.RIGHT));
        writer.append("</oddHeader>");
        writer.append("<oddFooter>");
        if (footer.get(Position.LEFT) != null) writer.append(footer.get(Position.LEFT));
        if (footer.get(Position.CENTER) != null) writer.append(footer.get(Position.CENTER));
        if (footer.get(Position.RIGHT) != null) writer.append(footer.get(Position.RIGHT));
        writer.append("</oddFooter></headerFooter>");


        if(!comments.isEmpty()) {
            writer.append("<drawing r:id=\"d\"/>");
            writer.append("<legacyDrawing r:id=\"v\"/>");
        }
        writer.append("</worksheet>");
        workbook.endFile();

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
     * <li>When a {@link Worksheet} is flushed, no other worksheet can be flushed until {@link #finish()} is called.</li>
     * </ul>
     *
     * @throws IOException If an I/O error occurs.
     */
    public void flush() throws IOException {
        if(writer == null) {
            int index = workbook.getIndex(this);
            writer = workbook.beginFile("xl/worksheets/sheet" + index + ".xml");
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            writer.append("<sheetPr filterMode=\"" + "false" + "\"><pageSetUpPr fitToPage=\"" + fitToPage + "\" autoPageBreaks=\"" + autoPageBreaks + "\"/></sheetPr>");
            writer.append("<dimension ref=\"A1\"/>");
            writer.append("<sheetViews><sheetView workbookViewId=\"0\"");
            if(!showGridLines) {
                writer.append(" showGridLines=\"false\"");
            }
            if(zoomScale != 100) {
                writer.append(" zoomScale=\"").append(zoomScale).append("\"");
            }
            writer.append(">");
            if(freezeLeftColumns > 0 || freezeTopRows > 0) {
                writeFreezePane(writer);
            }
            writer.append("</sheetView>");
            writer.append("</sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/>");
            int nbCols = rows.stream().filter(Objects::nonNull).map(r -> r.length).reduce(0, Math::max);
            if (nbCols > 0) {
                writeCols(writer, nbCols);
            }
            writer.append("<sheetData>");
        }

        for (int r = flushedRows; r < rows.size(); ++r) {
            Cell[] row = rows.get(r);
            if (row != null) {
                writeRow(writer, r, hiddenRows.contains(r),
                		rowHeights.get(r), row);
            }
            rows.set(r, null); // free flushed row data
        }
        flushedRows = rows.size() - 1;


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
     * @param customHeight Whether custom row height is set
     * @param rowHeight Row height value in points to be set if customHeight is true
     * @param row Cells in the row.
     * @throws IOException If an I/O error occurs.
     */
    private static void writeRow(Writer w, int r, boolean isHidden,
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
        w.append(">");
        for (int c = 0; c < row.length; ++c) {
            if (row[c] != null) {
                row[c].write(w, r, c);
            }
        }
        w.append("</row>");
    }

    /**
     * Assign a note/comment to a cell.
     * The comment popup will be twice the size of the cell and will be initially hidden.
     * <p>
     * Comments are stored in memory till call to {@link #finish()} - calling {@link #flush()} does not write them to output stream.
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
        this.fitToPage = true;
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
     * Gets input text form and converts it into what will be written in the
     * <header>/<footer> xml tag
     * @param text Header/footer text input form
     * @return (partial) String content of <header> or <footer> XML tag
     */
    private String prepareForXml(String text) {
        switch(text.toLowerCase()) {
            case "page 1 of ?":
                return "Page &amp;P of &amp;N";
            case "page 1, sheetname":
                return "Page &amp;P, &amp;A";
            case "page 1":
                return "Page &amp;P";
            case "sheetname":
                return "&amp;A";
            default:
                return text;
        }
    }

    /**
     * Set footer text.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     */
    public void footer(String text, Position position) {
        this.footer.put(position, "&amp;" + position.getPos() +
                                  prepareForXml(text));
    }

    /**
     * Set footer text with specified font size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontSize - integer describing font size
     */
    public void footer(String text, Position position, int fontSize) {
        this.footer.put(position, "&amp;" + position.getPos() +
                                  "&amp;&quot;Times New Roman,Regular&quot;&amp;" + fontSize +
                                  "&amp;K000000" + prepareForXml(text));
    }

    /**
     * Set footer text with specified font and size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontName - font name (e.g., "Arial")
     * @param fontSize - integer describing font size
     */
    public void footer(String text, Position position, String fontName, int fontSize) {
        this.footer.put(position, "&amp;" + position.getPos() +
                                  "&amp;&quot;" + fontName + ",Regular&quot;&amp;" + fontSize +
                                  "&amp;K000000" + prepareForXml(text));
    }

    /**
     * Set header text.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontName - font name (e.g., "Arial")
     * @param fontSize - integer describing font size
     */
    public void header(String text, Position position, String fontName, int fontSize) {
        this.header.put(position, "&amp;" + position.getPos() +
                                  "&amp;&quot;" + fontName + ",Regular&quot;&amp;" + fontSize +
                                  "&amp;K000000" + prepareForXml(text));
    }

    /**
     * Set header text with specified font size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontSize - integer describing font size
     */
    public void header(String text, Position position, int fontSize) {
        this.header.put(position, "&amp;" + position.getPos() +
                                  "&amp;&quot;Times New Roman,Regular&quot;&amp;" + fontSize +
                                  "&amp;K000000" + prepareForXml(text));
    }

    /**
     * Set header text with specified font and size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     */
    public void header(String text, Position position) {
        this.header.put(position, "&amp;" + position.getPos() +
                                  prepareForXml(text));
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
}
