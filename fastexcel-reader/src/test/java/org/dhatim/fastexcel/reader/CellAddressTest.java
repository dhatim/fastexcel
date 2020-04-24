package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

class CellAddressTest {
    @Test
    void refs() {
        assertCellAddress("A1", 0, 0);
        assertCellAddress("$A1", 0, 0);
        assertCellAddress("A$1", 0, 0);
        assertCellAddress("$A$1", 0, 0);
        assertCellAddress("B2", 1, 1);
        assertCellAddress("AA10", 9, 26);
        assertCellAddress("$AA$10", 9, 26);
        assertCellAddress("CW1", 0, 100);
        // max Excel address:
        assertCellAddress("XFD1048576", 1_048_576 - 1, 16_384 - 1);
        assertCellAddress("$XFD1048576", 1_048_576 - 1, 16_384 - 1);
        assertCellAddress("$XFD$1048576", 1_048_576 - 1, 16_384 - 1);
        // more then max
        assertCellAddress("ZZZ9999999", 9999998, 18277);
    }

    private void assertCellAddress(String ref, int row, int col) {
        assertCellAddressFromRef(ref, row, col);
        assertCellAddressFromRef(ref.toLowerCase(), row, col);
        assertEquals(ref.replaceAll("\\$", ""), new CellAddress(row, col).toString());
    }

    private void assertCellAddressFromRef(String ref, int row, int col) {
        CellAddress fromRef = new CellAddress(ref);
        assertEquals(row, fromRef.getRow(), "row: " + ref);
        assertEquals(col, fromRef.getColumn(), "col: " + ref);
    }

}
