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

import java.util.Locale;
import java.util.Objects;

public final class CellAddress implements Comparable<CellAddress> {

    public static final CellAddress A1 = new CellAddress(0, 0);

    private static final char ABSOLUTE_REFERENCE_MARKER = '$';

    private final int row;
    private final int col;

    public CellAddress(int row, int column) {
        this.row = row;
        this.col = column;
    }

    public CellAddress(String address) {
        int length = address.length();

        int loc = 0;
        // step over column name chars until first digit for row number.
        for (; loc < length; loc++) {
            char ch = address.charAt(loc);
            if (Character.isDigit(ch)) {
                break;
            }
        }

        String sCol = address.substring(0, loc).toUpperCase(Locale.ROOT);
        String sRow = address.substring(loc);

        this.row = Integer.parseInt(sRow) - 1;
        this.col = convertColStringToIndex(sCol);
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
        if (r != 0) {
            return r;
        }
        return 0;
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

    private static String convertNumToColString(int col) {
        // Excel counts column A as the 1st column, we
        // treat it as the 0th one
        int excelColNum = col + 1;

        StringBuilder colRef = new StringBuilder(2);
        int colRemain = excelColNum;

        while (colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }
            colRemain = (colRemain - thisPart) / 26;

            // The letter A is at 65
            char colChar = (char) (thisPart + 64);
            colRef.insert(0, colChar);
        }

        return colRef.toString();
    }

    private static int convertColStringToIndex(String ref) {
        int retval = 0;
        char[] refs = ref.toUpperCase(Locale.ROOT).toCharArray();
        for (int k = 0; k < refs.length; k++) {
            char thechar = refs[k];
            if (thechar == ABSOLUTE_REFERENCE_MARKER) {
                if (k != 0) {
                    throw new IllegalArgumentException("Bad col ref format '" + ref + "'");
                }
                continue;
            }

            // Character is uppercase letter, find relative value to A
            retval = (retval * 26) + (thechar - 'A' + 1);
        }
        return retval - 1;
    }

}
