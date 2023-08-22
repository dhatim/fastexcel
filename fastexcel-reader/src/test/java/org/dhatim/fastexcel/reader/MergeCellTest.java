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

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Stream;

class MergeCellTest {
        @org.junit.jupiter.api.Test
    void test() throws IOException {
        try (InputStream is = Resources.open("/xlsx/merge_cells.xlsx"); ReadableWorkbook wb = new ReadableWorkbook(is)) {
            Sheet sheet = wb.getFirstSheet();
            Map<Integer, Row> rowMap = new HashMap<>();
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(r -> rowMap.put(r.getRowNum(), r));
            }

            Row r1 = rowMap.get(1);
            assertTrue(r1.getCell(0).isMerged());
            assertEquals(0, r1.getCell(0).getMergedCellAddress().getRow());
            assertEquals(0, r1.getCell(0).getMergedCellAddress().getColumn());
            assertTrue(r1.getCell(1).isMerged());
            assertEquals(0, r1.getCell(1).getMergedCellAddress().getRow());
            assertEquals(0, r1.getCell(1).getMergedCellAddress().getColumn());
            assertTrue(r1.getCell(2).isMerged());
            assertEquals(0, r1.getCell(2).getMergedCellAddress().getRow());
            assertEquals(2, r1.getCell(2).getMergedCellAddress().getColumn());
            Row r2 = rowMap.get(2);
            assertTrue(r2.getCell(0).isMerged());
            assertEquals(0, r2.getCell(0).getMergedCellAddress().getRow());
            assertEquals(0, r2.getCell(0).getMergedCellAddress().getColumn());
            assertTrue(r2.getCell(1).isMerged());
            assertEquals(0, r2.getCell(1).getMergedCellAddress().getRow());
            assertEquals(0, r2.getCell(1).getMergedCellAddress().getColumn());
            assertTrue(r2.getCell(2).isMerged());
            assertEquals(0, r2.getCell(2).getMergedCellAddress().getRow());
            assertEquals(2, r2.getCell(2).getMergedCellAddress().getColumn());
            Row r3 = rowMap.get(3);
            assertTrue(r3.getCell(0).isMerged());
            assertEquals(2, r3.getCell(0).getMergedCellAddress().getRow());
            assertEquals(0, r3.getCell(0).getMergedCellAddress().getColumn());
            assertTrue(r3.getCell(1).isMerged());
            assertEquals(2, r3.getCell(1).getMergedCellAddress().getRow());
            assertEquals(0, r3.getCell(1).getMergedCellAddress().getColumn());
            assertFalse(r3.getCell(2).isMerged());
            assertNull(r3.getCell(2).getMergedCellAddress());
        }
    }
}
