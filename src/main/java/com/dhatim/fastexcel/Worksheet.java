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
package com.dhatim.fastexcel;

import java.io.IOException;
import java.time.ZonedDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/**
 * A worksheet is a set of cells.
 */
public class Worksheet {

    private final Workbook workbook;
    private final String name;
    /**
     * List of rows. A row is an array of cells.
     */
    private final List<Cell[]> rows = new ArrayList<>();
    /**
     * Ranges of merged cells.
     */
    private final Set<Range> mergedRanges = new HashSet<>();
    /**
     * List of ranges where shading to alternate rows is defined.
     */
    private final List<AlternateShading> alternateShadingRanges = new ArrayList<>();
    /**
     * Is this worksheet construction completed?
     */
    private boolean finished = false;

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
        // Add null for missing rows.
        while (r >= rows.size()) {
            rows.add(null);
        }
        Cell[] row = rows.get(r);
        if (row == null) {
            int columns = Math.max(c + 1, (r > 0 && rows.get(r - 1) != null) ? rows.get(r - 1).length : c + 1);
            row = new Cell[columns];
            rows.set(r, row);
        } else if (c >= row.length) {
            int columns = Math.max(c + 1, (r > 0 && rows.get(r - 1) != null) ? rows.get(r - 1).length : c + 1);
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

    /**
     * Set the cell value at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param value Cell value. Supported types are
     * {@link String}, {@link Date}, {@link ZonedDateTime}. Note Excel
     * timestamps do not carry any timezone information; {@link Date} values are
     * converted to an Excel serial number with the system timezone. If you need
     * a specific timezone, prefer passing a {@link ZonedDateTime}. and
     * {@link Number} implementations.
     */
    public void value(int r, int c, Object value) {
        if (value instanceof String) {
            cell(r, c).value = new CachedString(workbook.cacheString((String) value));
        } else if (value == null || value instanceof Number) {
            cell(r, c).value = value;
        } else if (value instanceof Date) {
            cell(r, c).value = TimestampUtil.convertDate((Date) value);
        } else if (value instanceof ZonedDateTime) {
            cell(r, c).value = TimestampUtil.convertZonedDateTime((ZonedDateTime) value);
        } else {
            throw new IllegalArgumentException("No supported cell type for " + value.getClass().toString());
        }
    }

    /**
     * Get the cell value (or formula) at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @return Cell value (or {@link Formula}).
     */
    public Object value(int r, int c) {
        Cell[] row = r < rows.size() ? rows.get(r) : null;
        Cell cell = row == null || c >= row.length ? null : row[c];
        Object o = cell == null ? null : cell.value;
        if (o instanceof CachedString) {
            o = workbook.getString(((CachedString) o).getIndex());
        }
        return o;
    }

    /**
     * Set the cell formula at the given coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param expression Cell formula expression.
     */
    public void formula(int r, int c, String expression) {
        cell(r, c).value = new Formula(expression);
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
     * Write column definitions of this worksheet as an XML element.
     *
     * @param w Output writer.
     * @param nbCols Number of columns.
     * @throws IOException If an I/O error occurs.
     */
    private void writeCols(Writer w, int nbCols) throws IOException {
        // Adjust column widths
        w.append("<cols autoFit=\"true\">");
        for (int c = 0; c < nbCols; ++c) {
            double maxWidth = 0;
            for (int r = 0; r < rows.size(); ++r) {
                // Exclude merged cells from computation
                boolean ignore = false;
                for (Range range : mergedRanges) {
                    if (range.contains(r, c)) {
                        ignore = true;
                        break;
                    }
                }
                Object o = ignore ? null : value(r, c);
                if (o != null && !(o instanceof Formula)) {
                    int length = o.toString().length();
                    maxWidth = Math.max(maxWidth, (int) ((length * 7 + 10) / 7.0 * 256) / 256.0);
                }
            }
            if (maxWidth > 0) {
                w.append("<col min=\"").append(c + 1).append("\" max=\"").append(c + 1).append("\" width=\"").append(Math.min(255.0, maxWidth)).append("\" customWidth=\"true\" bestFit=\"true\"/>");
            }
        }
        w.append("</cols>");
    }

    /**
     * Finish the construction of this worksheet.
     *
     * @throws IOException If an I/O error occurs.
     */
    public void finish() throws IOException {
        if (finished) {
            throw new IllegalStateException("Worksheet already finished.");
        }
        int index = workbook.getIndex(this);
        workbook.writeFile("xl/worksheets/sheet" + Integer.toString(index) + ".xml", w -> {
            w.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/>");
            int nbCols = rows.stream().filter(r -> r != null).map(r -> r.length).reduce(0, Math::max);
            if (nbCols > 0) {
                writeCols(w, nbCols);
            }
            w.append("<sheetData>");
            for (int r = 0; r < rows.size(); ++r) {
                Cell[] row = rows.get(r);
                if (row != null) {
                    writeRow(w, r, row);
                }
            }
            w.append("</sheetData>");
            if (!mergedRanges.isEmpty()) {
                w.append("<mergeCells>");
                for (Range r : mergedRanges) {
                    w.append("<mergeCell ref=\"").append(r.toString()).append("\"/>");
                }
                w.append("</mergeCells>");
            }
            for (AlternateShading a : alternateShadingRanges) {
                a.write(w);
            }
            w.append("<pageMargins bottom=\"0.75\" footer=\"0.3\" header=\"0.3\" left=\"0.7\" right=\"0.7\" top=\"0.75\"/></worksheet>");
        });

        // Free memory; we no longer need this data
        rows.clear();
        finished = true;
    }

    /**
     * Write a row as an XML element.
     *
     * @param w Output writer.
     * @param r Zero-based row number.
     * @param row Cells in the row.
     * @throws IOException If an I/O error occurs.
     */
    private static void writeRow(Writer w, int r, Cell... row) throws IOException {
        w.append("<row r=\"").append(r + 1).append("\">");
        for (int c = 0; c < row.length; ++c) {
            if (row[c] != null) {
                row[c].write(w, r, c);
            }
        }
        w.append("</row>");
    }
}
