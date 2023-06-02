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

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Stream;

public class Row implements Iterable<Cell> {

    private final int rowNum;
    private final List<Cell> cells;
    private final int physicalCellCount;

    Row(int rowNum, int physicalCellCount, List<Cell> cells) {
        this.rowNum = rowNum;
        this.physicalCellCount = physicalCellCount;
        this.cells = cells;
    }

    /**
     * Returns a cell in this row by column index;
     * @param index - zero-based column index
     * @return Cell value
     * @throws IndexOutOfBoundsException if index is invalid
     */
    public Cell getCell(int index) {
        return cells.get(index);
    }

    public Cell getCell(CellAddress address) {
        if (rowNum -1 != address.getRow()) {
            throw new IllegalArgumentException("The given address " + address + " concerns another row (" + rowNum + ")");
        }
        return getCell(address.getColumn());
    }

    public List<Cell> getCells(int beginIndex, int endIndex) {
        return cells.subList(beginIndex, endIndex);
    }

    public Optional<Cell> getOptionalCell(int index) {
        return index < 0 || index >= cells.size() ? Optional.empty() : Optional.ofNullable(cells.get(index));
    }

    public Optional<Cell> getFirstNonEmptyCell() {
        return stream().filter(Objects::nonNull).filter(cell -> !cell.getText().isEmpty()).findFirst();
    }

    public int getCellCount() {
        return cells.size();
    }

    public boolean hasCell(int index) {
        return index >= 0 && index < cells.size() && cells.get(index) != null;
    }

    /**
     * Get row number of this row
     * @return the row number (1 based)
     */
    public int getRowNum() {
        return rowNum;
    }

    public int getPhysicalCellCount() {
        return physicalCellCount;
    }

    @Override
    public String toString() {
        return "Row " + rowNum + ' ' + cells;
    }

    @Override
    public Iterator<Cell> iterator() {
        return cells.iterator();
    }

    public Stream<Cell> stream() {
        return cells.stream();
    }

    public Optional<String> getCellAsString(int cellIndex) {
        return getOptionalCell(cellIndex).map(Cell::asString);
    }

    public Optional<LocalDateTime> getCellAsDate(int cellIndex) {
        return getOptionalCell(cellIndex).map(Cell::asDate);
    }

    public Optional<BigDecimal> getCellAsNumber(int cellIndex) {
        return getOptionalCell(cellIndex).map(Cell::asNumber);
    }

    public Optional<Boolean> getCellAsBoolean(int cellIndex) {
        return getOptionalCell(cellIndex).map(Cell::asBoolean);
    }

    public String getCellText(int cellIndex) {
        return getOptionalCell(cellIndex).map(Cell::getText).orElse("");
    }

    public Optional<String> getCellRawValue(int cellIndex) {
        return getOptionalCell(cellIndex).map(Cell::getRawValue);
    }

}
