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
import java.util.stream.IntStream;

/**
 * A worksheet is a set of cells.
 */
public class Worksheet extends AbstractWorksheet {

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
    private final DynamicByteArray groupColumns = new DynamicByteArray(MAX_COLS);
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
     * Range of row where will be inserted auto filter
     */
    private Range autoFilterRange = null;

    /**
     * List of named ranges.
     */
    private Map<String, Range> namedRanges = new LinkedHashMap<>();

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
        super(workbook, name);
    }

    /**
     * Get repeating rows defined for the print setup.
     *
     * @return List representing a range of rows to be repeated
     *              on each page when printing.
     */
    public RepeatRowRange getRepeatingRows() {
        return repeatingRows;
    }

    /**
     * Get cell range that autofilter is applied to.
     *
     * @return Range of cells that autofilter is set to
     *             (null if autofilter is not set).
     */
    public Range getAutoFilterRange() {
        return autoFilterRange;
    }

    /**
     * Get repeating cols defined for the print setup.
     *
     * @return List representing a range of columns to be repeated
     *              on each page when printing.
     */
    public RepeatColRange getRepeatingCols() {
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
            int groupLevel = groupColumns.get(c);
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
     * Finish the construction of this worksheet. This creates the worksheet
     * file on the workbook's output stream. Rows and cells in this worksheet
     * are then destroyed.
     *
     * @throws IOException If an I/O error occurs.
     */
    @Override
    public void finish() throws IOException {

        if (isFinished()) {
            return;
        }

        initDocumentAndFlush();

        int index = workbook.getIndex(this);
        setupSheetPassword(writer);

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

        setupDataValidations();
        setupHyperlinkRanges();

        setupPageMargins();
        setupPage();
        setupFooter();
        setupComments();
        setupTables();

        writer.append("</worksheet>");
        workbook.endFile();

        writeComments(index);
        writeTables();
        writeRelationships(index);

        // Free memory; we no longer need this data
        rows.clear();
        finished = true;
    }

    /**
     * Write all the rows currently in memory to the workbook's output stream.
     * Call this method periodically when working with huge data sets.
     * After calling {@link #initDocumentAndFlush()}, all the rows created so far become inaccessible.<br>
     * Notes:<br>
     * <ul>
     * <li>All columns must be defined before calling this method:
     * do not add or merge columns after calling {@link #initDocumentAndFlush()}.</li>
     * <li>When a {@link Worksheet} is flushed, no other worksheet can be flushed until {@link #close()} (or  the old fashion way {@link #finish()}) is called.</li>
     * </ul>
     *
     * @throws IOException If an I/O error occurs.
     */
    @Override
    public void initDocumentAndFlush() throws IOException {
        if (writer == null) {
            writer = initWriter();
        }

        initDocumentTags();

        int nbCols = rows.stream().filter(Objects::nonNull).mapToInt(r -> r.length).max().orElse(0);
        int maxHideCol = hiddenColumns.stream().mapToInt(a -> a).max().orElse(0);
        int maxStyleCol = colStyles.values().stream().mapToInt(Column::getColNumber).max().orElse(0);
        int maxNoZeroIndex = groupColumns.getMaxNoZeroIndex();

        if (nbCols > 0 || !hiddenColumns.isEmpty()||maxNoZeroIndex!=-1 || !colStyles.isEmpty()) {
            int maxCol = Math.max(nbCols, Math.max(Math.max(maxHideCol,maxNoZeroIndex), maxStyleCol) + 1);
            writeCols(writer, maxCol);
        }

        appendRows();
        writer.flush();
    }

    public void appendRows() throws IOException {
        writer.append("<sheetData>");

        int nbRows = rows.size();
        int maxHideRow = hiddenRows.stream().mapToInt(a -> a).max().orElse(0);
        int maxGroupRow = groupRows.getMaxNoZeroIndex();
        int maxRow = Math.max(nbRows, Math.max(maxGroupRow,maxHideRow) + 1);

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

        writer.append("</sheetData>");
        flushedRows = maxRow - 1;
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

    public void groupCols(int from , int to) {
        IntStream.rangeClosed(Math.min(from,to),Math.max(from,to)).forEach(groupColumns::increase);
    }

    public void groupRows(int from , int to) {
        IntStream.rangeClosed(Math.min(from,to),Math.max(from,to)).forEach(groupRows::increase);
    }
}
