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
 *
 */
package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * @author gamerover98
 */
public class StreamWorksheet extends AbstractWorksheet {

    private int rows;
    private Cell[] cells;

    /**
     * Constructor.
     *
     * @param workbook Parent workbook.
     * @param name     Worksheet name.
     */
    public StreamWorksheet(Workbook workbook, String name) {
        super(workbook, name);
    }

    @Override
    public void initDocumentAndFlush() {
        if (writer == null) {
            throw new IllegalStateException("The stream is not started.");
        }
    }

    @Override
    public void finish() throws IOException {
        if (isFinished()) {
            return;
        }

        writer.append("</sheetData>");

        int index = workbook.getIndex(this);
        setupSheetPassword(writer);

        setupDataValidations();
        setupHyperlinkRanges();

        setupPageMargins();
        setupPage();
        setupFooter();
        setupComments();
        setupTables();

        writer.append("</worksheet>");
        workbook.endFile();

        writeComments(index);
        writeTables();
        writeRelationships(index);

        finished = true;
    }

    public StreamWorksheet start(int columns) throws IOException {

        int index = workbook.getIndex(this);
        writer = workbook.beginFile("xl/worksheets/sheet" + index + ".xml");

        initDocumentTags();
        initColumns(columns);

        this.cells = new Cell[columns];

        for (int i = 0; i < columns; i++) {
            cells[i] = new Cell();
        }

        writer.append("<sheetData>");
        return this;
    }

    public StreamWorksheet appendRow(CellSupplier cellSupplier) throws IOException {
        int index = 0;

        for (Cell cell : cells) {
            cell.setValue(workbook, String.valueOf(cellSupplier.get(index++)));
        }

        writeRow(writer, rows++, false, (byte) 0, null, cells);
        return this;
    }

    protected void initColumns(int columns) throws IOException {

        if (columns <= 0) {
            throw new IllegalArgumentException("The columns must be greater than 0.");
        }

        writer.append("<cols>");

        for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
            writeCol(writer,
                    columnIndex,
                    DEFAULT_COL_WIDTH,
                    true,
                    false,
                    0,
                    Column.noStyle(this, columnIndex).getStyle());
        }

        writer.append("</cols>");
    }

    public interface CellSupplier {

        Object get(int column);
    }
}
