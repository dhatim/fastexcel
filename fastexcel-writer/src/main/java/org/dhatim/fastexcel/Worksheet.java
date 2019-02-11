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
     * List of DataValidations for this worksheet
     */
    private final List<DataValidation> dataValidations = new ArrayList<>();
    /**
     * List of ranges where shading to alternate rows is defined.
     */
    private final List<AlternateShading> alternateShadingRanges = new ArrayList<>();
    /**
     * List of rows to hide
     */
    private final Set<Integer> hiddenRows = new HashSet<>();

    /**
     * List of columns to hide
     */
    private final Set<Integer> hiddenColumns = new HashSet<>();

    /**
     * Map of columns and they're widths
     */
    private final Map<Integer, Double> colWidths = new HashMap<>();
    /**
     * Is this worksheet construction completed?
     */
    private boolean finished;

    /**
     * The visibility state of this sheet.
     */
    private VisibilityState visibilityState;

    /**
     * The hashed password that protects this sheet.
     */
    private String passwordHash;

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
        alternateShadingRanges.add(new AlternateShading(range, getWorkbook().cacheAlternateShadingFillColor(fill)));
    }

    void addValidation(DataValidation validation) {
        dataValidations.add(validation);
    }

    /**
     * Sets the visibility state of the sheet
     * <p>
     * This is done by setting the {@code state} attribute in the workbook.xml.
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
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value. Supported types are
     * {@link String}, {@link Date}, {@link LocalDate}, {@link LocalDateTime}, {@link ZonedDateTime},
     * {@link Number} and {@link Boolean} implementations. Note Excel timestamps do not carry
     * any timezone information; {@link Date} values are converted to an Excel
     * serial number with the system timezone. If you need a specific timezone,
     * prefer passing a {@link ZonedDateTime}.
     */
    public void value(int r, int c, Object value) {
        cell(r, c).setValue(workbook, value);
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
        return mergedRanges.stream().filter(range -> range.contains(r, c)).findAny().isPresent();
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
        if (!mergedRanges.isEmpty()) {
            writer.append("<mergeCells>");
            for (Range r : mergedRanges) {
                writer.append("<mergeCell ref=\"").append(r.toString()).append("\"/>");
            }
            writer.append("</mergeCells>");
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

        if (passwordHash != null) {
            writer.append("<sheetProtection password=\"").append(passwordHash).append("\" ");
            for (SheetProtectionOption option : SheetProtectionOption.values()) {
                if (option.getDefaultValue() != sheetProtectionOptions.contains(option)) {
                    writer.append(option.getName()).append("=\"").append(Boolean.toString(!option.getDefaultValue())).append("\" ");
                }
            }
            writer.append("/>");
        }

        writer.append("<pageMargins bottom=\"0.75\" footer=\"0.3\" header=\"0.3\" left=\"0.7\" right=\"0.7\" top=\"0.75\"/></worksheet>");
        workbook.endFile();

        // Free memory; we no longer need this data
        rows.clear();
        finished = true;
    }

    /**
     * Write all the rows currently in memory to the workbook's output stream.
     * Call this method periodically when working with huge data sets.
     * After calling {@link #flush()}, all the rows created so far become inaccessible.<br/>
     * Notes:<br/>
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
            writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/>");
            int nbCols = rows.stream().filter(r -> r != null).map(r -> r.length).reduce(0, Math::max);
            if (nbCols > 0) {
                writeCols(writer, nbCols);
            }
            writer.append("<sheetData>");
        }
        for (int r = flushedRows; r < rows.size(); ++r) {
            Cell[] row = rows.get(r);
            if (row != null) {
                writeRow(writer, r, hiddenRows.contains(r), row);
            }
            rows.set(r, null); // free flushed row data
        }
        flushedRows = rows.size() - 1;
        writer.flush();
    }

    /**
     * Write a row as an XML element.
     *
     * @param w Output writer.
     * @param r Zero-based row number.
     * @param isHidden Whether or not this row is hidden
     * @param row Cells in the row.
     * @throws IOException If an I/O error occurs.
     */
    private static void writeRow(Writer w, int r, boolean isHidden, Cell... row) throws IOException {
        w.append("<row r=\"").append(r + 1);
        if (isHidden) {
            w.append("\" hidden=\"true");
        }
        w.append("\">");
        for (int c = 0; c < row.length; ++c) {
            if (row[c] != null) {
                row[c].write(w, r, c);
            }
        }
        w.append("</row>");
    }
}
