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

import java.util.*;
import java.util.stream.IntStream;

/**
 * Definition of a range of cells.
 */
public class Range implements Ref {

    /**
     * Worksheet where this range is defined.
     */
    private final Worksheet worksheet;
    /**
     * Top row.
     */
    private final int top;
    /**
     * Left column.
     */
    private final int left;
    /**
     * Bottom row.
     */
    private final int bottom;
    /**
     * Right column.
     */
    private final int right;

    /**
     * Constructor. Note coordinates are reordered if necessary to make sure
     * {@code top} &lt;= {@code bottom} and {@code left} &lt;= {@code right}.
     *
     * @param worksheet Parent worksheet.
     * @param top Top row.
     * @param left Left column.
     * @param bottom Bottom row.
     * @param right Right column.
     */
    Range(Worksheet worksheet, int top, int left, int bottom, int right) {
        this.worksheet = Objects.requireNonNull(worksheet);

        // Check limits
        if (top < 0 || top >= Worksheet.MAX_ROWS || bottom < 0 || bottom >= Worksheet.MAX_ROWS) {
            throw new IllegalArgumentException();
        }
        if (left < 0 || left >= Worksheet.MAX_COLS || right < 0 || right >= Worksheet.MAX_COLS) {
            throw new IllegalArgumentException();
        }
        this.top = Math.min(top, bottom);
        this.left = Math.min(left, right);
        this.bottom = Math.max(bottom, top);
        this.right = Math.max(right, left);
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
     * Get top row.
     *
     * @return Top row.
     */
    public int getTop() {
        return top;
    }

    /**
     * Get left column.
     *
     * @return Left column.
     */
    public int getLeft() {
        return left;
    }

    /**
     * Get bottom row.
     *
     * @return Bottom row.
     */
    public int getBottom() {
        return bottom;
    }

    /**
     * Get right column.
     *
     * @return Right column.
     */
    public int getRight() {
        return right;
    }

    @Override
    public int hashCode() {
        return Objects.hash(worksheet, top, left, bottom, right);
    }

    @Override
    public boolean equals(Object obj) {
        boolean result;
        if (obj != null && obj.getClass() == this.getClass()) {
            Range other = (Range) obj;
            result = Objects.equals(worksheet, other.worksheet) && Objects.equals(top, other.top) && Objects.equals(left, other.left) && Objects.equals(bottom, other.bottom) && Objects.equals(right, other.right);
        } else {
            result = false;
        }
        return result;
    }

    @Override
    public String toString() {
        return colToString(left) + (top + 1) + ':' + colToString(right) + (bottom + 1);
    }

    /**
     * Get an absolute reference to this Range.
     *
     * ex: $A$1:$A$5
     *
     * @return absolute reference
     */
    public String toAbsoluteString() {
        return '$' + colToString(left) + '$' + (top + 1) + ":$" + colToString(right) + '$' + (bottom + 1);
    }

    /**
     * Get a new style setter for this range.
     *
     * @return Newly created style setter.
     */
    public StyleSetter style() {
        return new StyleSetter(this);
    }

    /**
     * Merge cells within this range.
     */
    public void merge() {
        worksheet.merge(this);
    }

    /**
     * Check if this range contains the given cell coordinates.
     *
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @return {@code true} if this range contains the given cell coordinates.
     */
    public boolean contains(int r, int c) {
        return r >= top && r <= bottom && c >= left && c <= right;
    }

    /**
     * Apply shading to alternate rows in this range with the given fill
     * pattern.
     *
     * @param fill Fill pattern.
     */
    void shadeAlternateRows(Fill fill) {
        worksheet.shadeAlternateRows(this, fill);
    }

    void shadeRows(Fill fill, int eachNRows) {
        worksheet.shadeRows(this, fill, eachNRows);
    }

    /**
     * Construct a new ListDataValidation
     *
     * @param listRange The Range of the list this validation references
     * @return a new list data validation object
     */
    public ListDataValidation validateWithList(Range listRange) {
        ListDataValidation listDataValidation = new ListDataValidation(this, listRange);
        worksheet.addValidation(listDataValidation);
        return listDataValidation;
    }

     /**
     * Construct a new ListDataValidation
     *
     * @param formula The custom validation formula
     * @return a new custom validation
     */
    public CustomDataValidation validateWithFormula(String formula) {
        CustomDataValidation customDataValidation = new CustomDataValidation(this, new Formula(formula));
        worksheet.addValidation(customDataValidation);
        return customDataValidation;
    }

    /**
     * Specifically define this range by assigning it a name.
     * It will be visible in the cell range dropdown menu.
     * 
     * @param name string representing the name of this cell range
     */
    public void setName(String name) {
        worksheet.addNamedRange(this, name);
    }

    /**
     * Return the set of styles used by the cells in this range.
     *
     * @return Set of styles.
     */
    Set<Integer> getStyles() {
        Set<Integer> result = new HashSet<>();
        for (int r = top; r <= bottom; ++r) {
            for (int c = left; c <= right; ++c) {
                result.add(getWorksheet().cell(r, c).getStyle());
            }
        }
        return result;
    }

    /**
     * Apply new (merged) styles to the cells in this range.
     *
     * @param styles Map giving new style for each old style.
     */
    void applyStyle(Map<Integer, Integer> styles) {
        for (int r = top; r <= bottom; ++r) {
            for (int c = left; c <= right; ++c) {
                Cell cell = getWorksheet().cell(r, c);
                cell.setStyle(styles.get(cell.getStyle()));
            }
        }
    }
    public void setHyperlink(HyperLink hyperLink){
        this.worksheet.value(top,left,hyperLink.getDisplayStr());
        this.worksheet.addHyperlink(this,hyperLink);
    }

    public Table createTable() {
        int columnCount = this.right - this.left + 1;
        String[] headers = IntStream.rangeClosed(1, columnCount).mapToObj(i -> "Column" + i).toArray(String[]::new);
        return createTable(headers);
    }

    public Table createTable(String... headers) {
        return worksheet.addTable(this, headers);
    }
}
