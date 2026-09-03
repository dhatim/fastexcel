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

import org.dhatim.fastexcel.common.CellAddressUtil;

final class CellAddress {
    private static final String[] CACHED_COLS = new String[1024];

    static {
        for (int i = 0; i < CACHED_COLS.length; i++) {
            CACHED_COLS[i] = convertNumToColStringImpl(i);
        }
    }

    static StringBuilder format(int row, int col) {
        return format(new StringBuilder(), row, col);
    }

    static StringBuilder format(StringBuilder sb, int row, int col) {
        sb.append(convertNumToColString(col));
        sb.append(row + 1);
        return sb;
    }

    static String convertNumToColString(int col) {
        if (col < CACHED_COLS.length) {
            return CACHED_COLS[col];
        }
        return convertNumToColStringImpl(col);
    }

    private static String convertNumToColStringImpl(int col) {
        return CellAddressUtil.convertNumToColString(col);
    }
}
