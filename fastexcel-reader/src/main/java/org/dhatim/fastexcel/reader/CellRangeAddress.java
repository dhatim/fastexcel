package org.dhatim.fastexcel.reader;

import java.util.Objects;

public final class CellRangeAddress {

    private final int firstRow;
    private final int lastRow;
    private final int firstCol;
    private final int lastCol;

    public CellRangeAddress(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;

        if (lastRow < firstRow || lastCol < firstCol) {
            throw new IllegalArgumentException("Invalid cell range, having lastRow < firstRow || lastCol < firstCol, " +
                    "had rows " + lastRow + " >= " + firstRow + " or cells " + lastCol + " >= " + firstCol);
        }
    }

    public int getFirstColumn() {
        return firstCol;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public int getLastColumn() {
        return lastCol;
    }

    public int getLastRow() {
        return lastRow;
    }

    public boolean isInRange(int row, int column) {
        return firstRow <= row && row <= lastRow && firstCol <= column && column <= lastCol;
    }

    public boolean isInRange(CellAddress cell) {
        return isInRange(cell.getRow(), cell.getColumn());
    }

    public boolean containsRow(int row) {
        return firstRow <= row && row <= lastRow;
    }

    public boolean containsColumn(int column) {
        return firstCol <= column && column <= lastCol;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj == this) {
            return true;
        }
        if (obj == null || obj.getClass() != getClass()) {
            return false;
        }
        CellRangeAddress other = (CellRangeAddress) obj;
        return firstRow == other.firstRow && firstCol == other.firstCol && lastRow == other.lastRow && lastCol == other.lastCol;
    }

    @Override
    public int hashCode() {
        return Objects.hash(firstRow, firstCol, lastRow, lastCol);
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        CellAddress.format(sb, firstRow, firstCol);
        sb.append(':');
        CellAddress.format(sb, lastRow, lastCol);
        return sb.toString();
    }

    public static CellRangeAddress valueOf(String ref) {
        int sep = ref.indexOf(':');
        CellAddress a;
        CellAddress b;
        if (sep == -1) {
            a = new CellAddress(ref);
            b = a;
        } else {
            a = new CellAddress(ref.substring(0, sep));
            b = new CellAddress(ref.substring(sep + 1));
        }
        return new CellRangeAddress(a.getRow(), b.getRow(), a.getColumn(), b.getColumn());
    }

}
