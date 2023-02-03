package org.dhatim.fastexcel;

import java.util.Comparator;
import java.util.Objects;
/**
 * This class is used to refer to the location of a cell.
 */
class Location implements Comparable<Location>, Ref {
    final int row;
    final int col;

    public Location(int row, int col) {
        this.row = row;
        this.col = col;
    }

    @Override
    public int compareTo(Location o) {
        return Comparator.comparingInt((Location location) -> location.row)
                .thenComparing(location1 -> location1.col)
                .compare(this, o);
    }

    @Override
    public boolean equals(Object obj) {
        if (obj instanceof Location) {
            Location that = (Location) obj;
            return this == that || (this.row == that.row && this.col == that.col);
        }
        return false;
    }

    @Override
    public int hashCode() {
        return Objects.hash(row,col);
    }

    @Override
    public String toString() {
        return colToString(col) + (row + 1);
    }
}
