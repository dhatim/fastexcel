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

import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Iterator;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;

class SimpleReaderTest {

    private static final Object[][] VALUES = {
            {1, "Lorem", date(2018, 1, 1), null, true},
            {2, "ipsum", date(2018, 1, 2), null, false},
            {3, "dolor", date(2018, 1, 3), null, true},
            {4, "sit", date(2018, 1, 4), null, false},
            {5, "amet", date(2018, 1, 5), null, true},
            {6, "consectetur", date(2018, 1, 6), null, false},
            {7, "adipiscing", date(2018, 1, 7), null, true},
            {8, "elit", date(2018, 1, 8), null, false},
            {9, "Ut", date(2018, 1, 9), null, true},
            {10, "nec", date(2018, 1, 10), null, false},
    };

    private static LocalDateTime date(int year, int month, int day) {
        return LocalDateTime.of(year, month, day, 0, 0);
    }

    @Test
    void test() throws IOException {
        try (InputStream is = Resources.open("/xlsx/simple.xlsx"); ReadableWorkbook wb = new ReadableWorkbook(is)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(r -> {
                    BigDecimal num = r.getCellAsNumber(0).orElse(null);
                    String str = r.getCellAsString(1).orElse(null);
                    LocalDateTime date = r.getCellAsDate(2).orElse(null);
                    Boolean bool = r.getCellAsBoolean(4).orElse(null);

                    Object[] values = VALUES[r.getRowNum() - 1];
                    assertThat(num).isEqualTo(BigDecimal.valueOf((Integer) values[0]));
                    assertThat(str).isEqualTo((String) values[1]);
                    assertThat(date).isEqualTo(values[2]);
                    assertThat(bool).isEqualTo(values[4]);
                });
            }
        }
    }

    @Test
    void testWithParseErrorOnNumber() throws IOException {
        try (InputStream is = Resources.open("/xlsx/parseError.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is, ReadingOptions.DEFAULT_READING_OPTIONS)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                Iterator<Row> it = rows.iterator();
                try {
                    it.hasNext();
                    fail("Must throw an exception");
                } catch (ExcelReaderException e) {
                    // OK
                }
            }
        }


        try (InputStream is = Resources.open("/xlsx/parseError.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is, new ReadingOptions(false, true))) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                Iterator<Row> it = rows.iterator();
                assertTrue(it.hasNext());
                Iterator<Cell> cellIt = it.next().iterator();
                assertTrue(cellIt.hasNext());
                Cell cell = cellIt.next();
                assertEquals(CellType.ERROR, cell.getType());
            }
        }
    }

    @Test
    public void testDefaultWorkbookPath() throws IOException {
        try (InputStream is = Resources.open("/xlsx/DefaultContentType.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is, new ReadingOptions(false, true))) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                Iterator<Row> it = rows.iterator();
                assertTrue(it.hasNext());
                Iterator<Cell> cellIt = it.next().iterator();
                assertTrue(cellIt.hasNext());
                Cell cell = cellIt.next();
                assertEquals(CellType.NUMBER, cell.getType());
                assertEquals(BigDecimal.ONE, cell.getValue());
            }
        }
    }

    @Test
    public void testDefaultWorkbookPath2() throws IOException {
        try (InputStream is = Resources.open("/xlsx/absolutePath.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is, new ReadingOptions(false, true))) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                Iterator<Row> it = rows.iterator();
                assertTrue(it.hasNext());
                Iterator<Cell> cellIt = it.next().iterator();
                assertTrue(cellIt.hasNext());
                Cell cell = cellIt.next();
                assertEquals(CellType.NUMBER, cell.getType());
                assertEquals(BigDecimal.ONE, cell.getValue());
            }
        }
    }

    @Test
    public void testCaseInsensitiveInFileNames() throws IOException {
        try (InputStream is = Resources.open("/xlsx/caseInsensitive.xlsx");
             ReadableWorkbook wb = new ReadableWorkbook(is, new ReadingOptions(false, true))) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                Iterator<Row> it = rows.iterator();
                assertTrue(it.hasNext());
                Iterator<Cell> cellIt = it.next().iterator();
                assertTrue(cellIt.hasNext());
                Cell cell = cellIt.next();
                assertEquals(CellType.STRING, cell.getType());
                assertEquals("A", cell.getValue());
            }
        }
    }
}
