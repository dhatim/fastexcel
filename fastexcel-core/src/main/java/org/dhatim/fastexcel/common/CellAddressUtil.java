package org.dhatim.fastexcel.common;

import java.nio.charset.StandardCharsets;

public final class CellAddressUtil {

    private static final int COL_RADIX = 'Z' - 'A' + 1;
    private static final int MAX_COL_CHARS = 3;

    private CellAddressUtil() {
    }

    public static String convertNumToColString(int col) {
        int excelColNum = col + 1;

        final byte[] colRef = new byte[MAX_COL_CHARS];
        int colRemain = excelColNum;
        int pos = MAX_COL_CHARS - 1;
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
}
