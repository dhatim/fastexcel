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
package org.dhatim.fastexcel.reader;

import java.nio.charset.StandardCharsets;
import java.util.Objects;

public final class CellAddress implements Comparable<CellAddress> {

    public static final CellAddress A1 = new CellAddress(0, 0);

    private static final char ABSOLUTE_REFERENCE_MARKER = '$';
    private static final int COL_RADIX = 'Z' - 'A' + 1;

    private final int row;
    private final int col;

    /**
     * Represents a cell address inside a sheet
     *
     * @param row    zero-based row index
     * @param column zero-based column index
     */
    public CellAddress(int row, int column) {
        this.row = row;
        this.col = column;
    }

    public CellAddress(String address) {
        final int length = address.length();
        if (length == 0) {
            this.row = 0;
            this.col = 0;
        } else {
            int offset = address.charAt(0) == ABSOLUTE_REFERENCE_MARKER ? 1 : 0;
            int col = 0;
            for (; offset < length; offset++) {
                final char c = address.charAt(offset);
                if (c == ABSOLUTE_REFERENCE_MARKER) {
                    offset++;
                    break; //next there must be digits
                }
                if (isAsciiDigit(c)) {
                    break;
                }
                col = col * COL_RADIX + toUpperCase(c) - (int) 'A' + 1;
            }
            this.col = col - 1;
            this.row = Integer.parseUnsignedInt(address.substring(offset)) - 1;
        }
    }

    public int getRow() {
        return row;
    }

    public int getColumn() {
        return col;
    }

    @Override
    public int compareTo(CellAddress other) {
        int r = row - other.row;
        if (r != 0) {
            return r;
        }
        r = col - other.col;
        return r;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj == this) {
            return true;
        }
        if (obj == null || obj.getClass() != getClass()) {
            return false;
        }
        CellAddress other = (CellAddress) obj;
        return row == other.row && col == other.col;
    }

    @Override
    public int hashCode() {
        return Objects.hash(row, col);
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        format(sb, row, col);
        return sb.toString();
    }

    static void format(StringBuilder sb, int row, int col) {
        sb.append(convertNumToColString(col));
        sb.append(row + 1);
    }

    public static String convertNumToColString(int col) {
        // Excel counts column A as the 1st column, we
        // treat it as the 0th one
        int excelColNum = col + 1;

        final int MAX_COL_CHARS = 3;
        final byte[] colRef = new byte[MAX_COL_CHARS];
        int colRemain = excelColNum;
        int pos = 2;
        while (colRemain > 0) {
            int thisPart = colRemain % COL_RADIX;
            if (thisPart == 0) {
                thisPart = COL_RADIX;
            }
            colRemain = (colRemain - thisPart) / COL_RADIX;

            colRef[pos--] = (byte) (thisPart + (int) 'A' - 1);
        }
        pos++;
        return new String(colRef, pos, (MAX_COL_CHARS - pos), StandardCharsets.ISO_8859_1);
    }


    private static final boolean isAsciiLowerCase(char c) {
        return 'a' <= c && c <= 'z';
    }

    private static final boolean isAsciiUpperCase(char c) {
        return 'A' <= c && c <= 'Z';
    }

    private static final boolean isAsciiDigit(char c) {
        return '0' <= c && c <= '9';
    }

    private final static char toUpperCase(char c) {
        if (isAsciiUpperCase(c)) {
            return c;
        }
        if (isAsciiLowerCase(c)) {
            return (char) (c + ('A' - 'a'));
        }
        throw new IllegalArgumentException("Unexpected char: " + c);
    }

}
