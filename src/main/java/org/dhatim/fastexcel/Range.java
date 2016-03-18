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

import java.util.HashSet;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Definition of a range of cells.
 */
public class Range {

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
        this.top = top <= bottom ? top : bottom;
        this.left = left <= right ? left : right;
        this.bottom = bottom >= top ? bottom : top;
        this.right = right >= left ? right : left;
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
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final Range other = (Range) obj;
        if (!Objects.equals(this.worksheet, other.worksheet)) {
            return false;
        }
        if (this.top != other.top) {
            return false;
        }
        if (this.left != other.left) {
            return false;
        }
        if (this.bottom != other.bottom) {
            return false;
        }
        return this.right == other.right;
    }

    /**
     * Convert a column index to a column name.
     *
     * @param c Zero-based column index.
     * @return Column name.
     */
    public static String colToString(int c) {
        StringBuilder sb = new StringBuilder();
        while (c >= 0) {
            sb.append((char) ('A' + (c % 26)));
            c = (c / 26) - 1;
        }
        return sb.reverse().toString();
    }

    @Override
    public String toString() {
        return colToString(left) + Integer.toString(top + 1) + ':' + colToString(right) + Integer.toString(bottom + 1);
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
}
