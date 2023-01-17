package org.dhatim.fastexcel;

import java.util.Comparator;

class Location implements Comparable<Location>, Ref {
    final int row;
    final int col;

    public Location(int row, int col) {
        this.row = row;
        this.col = col;
    }

    private int getRow() {
        return row;
    }

    private int getCol() {
        return col;
    }

    @Override
    public int compareTo(Location o) {
        return Comparator.comparingInt(Location::getRow)
                .thenComparing(Location::getCol)
                .compare(this, o);
    }

    @Override
    public String toString() {
        return colToString(col) + (row + 1);
    }

    @Override
    public boolean equals(Object obj) {
        if (obj != null) {
            if (obj instanceof Location) {
                return this.compareTo((Location)obj) == 0;
            }
        }
        return false;
    }
}
