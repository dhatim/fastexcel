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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.function.Consumer;
import org.apache.commons.io.output.NullOutputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import static org.junit.Assert.*;
import org.junit.Test;

public class Correctness {

    private byte[] writeWorkbook(Consumer<Workbook> consumer) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        Workbook wb = new Workbook(os, "Test", "1.0");
        consumer.accept(wb);
        wb.finish();
        return os.toByteArray();
    }

    @Test
    public void colToName() throws Exception {
        assertEquals("AA", Range.colToString(26));
        assertEquals("AAA", Range.colToString(702));
    }

    @Test(expected = IllegalArgumentException.class)
    public void noWorksheet() throws Exception {
        writeWorkbook(wb -> {
        });
    }

    @Test(expected = IllegalArgumentException.class)
    public void badVersion() throws Exception {
        Workbook dummy = new Workbook(new NullOutputStream(), "Test", "1.0.1");
    }

    @Test
    public void singleEmptyWorksheet() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1"));
    }

    @Test
    public void singleWorksheet() throws Exception {
        String sheetName = "Worksheet 1";
        String stringValue = "Sample text";
        Date dateValue = new Date();
        ZoneId timezone = ZoneId.of("Australia/Sydney");
        ZonedDateTime zonedDateValue = ZonedDateTime.ofInstant(dateValue.toInstant(), timezone);
        double doubleValue = 1.234;
        int intValue = 2_016;
        long longValue = 2_016_000_000_000L;
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet(sheetName);
            int i = 1;
            ws.value(i, i++, stringValue);
            ws.value(i, i++, dateValue);
            ws.value(i, i++, zonedDateValue);
            ws.value(i, i++, doubleValue);
            ws.value(i, i++, intValue);
            ws.value(i, i++, longValue);
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertEquals(0, xwb.getActiveSheetIndex());
        assertEquals(1, xwb.getNumberOfSheets());
        XSSFSheet xws = xwb.getSheet(sheetName);
        assertNull(xws.getRow(0));
        int i = 1;
        assertEquals(stringValue, xws.getRow(i).getCell(i++).getStringCellValue());
        assertEquals(dateValue, xws.getRow(i).getCell(i++).getDateCellValue());

        // Check zoned timestamp has the same textual representation as the Date extracted from the workbook
        // (Excel date serial numbers do not carry timezone information)
        assertEquals(DateTimeFormatter.ISO_LOCAL_DATE_TIME.format(zonedDateValue), DateTimeFormatter.ISO_LOCAL_DATE_TIME.format(ZonedDateTime.ofInstant(xws.getRow(i).getCell(i++).getDateCellValue().toInstant(), ZoneId.systemDefault())));
        assertEquals(doubleValue, xws.getRow(i).getCell(i++).getNumericCellValue(), 0);
        assertEquals(intValue, xws.getRow(i).getCell(i++).getNumericCellValue(), 0);
        assertEquals(longValue, xws.getRow(i).getCell(i++).getNumericCellValue(), 0);
    }

    @Test
    public void multipleWorksheets() throws Exception {
        int numWs = 10;
        int numRows = 5000;
        int numCols = 5;
        byte[] data = writeWorkbook(wb -> {
            CompletableFuture<Void>[] cfs = new CompletableFuture[numWs];
            for (int i = 0; i < cfs.length; ++i) {
                Worksheet ws = wb.newWorksheet("Sheet " + i);
                CompletableFuture<Void> cf = CompletableFuture.runAsync(() -> {
                    for (int j = 0; j < numCols; ++j) {
                        ws.value(0, j, "Column " + j);
                        ws.style(0, j).bold().fillColor(Color.GRAY2).set();
                        for (int k = 1; k <= numRows; ++k) {
                            switch (j) {
                                case 0:
                                    ws.value(k, j, "String value " + k);
                                    break;
                                case 1:
                                    ws.value(k, j, 2);
                                    break;
                                case 2:
                                    ws.value(k, j, 3L);
                                    break;
                                case 3:
                                    ws.value(k, j, 0.123);
                                    break;
                                case 4:
                                    ws.value(k, j, new Date());
                                    ws.style(k, j).format("yyyy-MM-dd HH:mm:ss").set();
                                    break;
                                default:
                                    throw new IllegalArgumentException();
                            }
                        }
                    }
                    ws.formula(numRows + 1, 1, "=SUM(" + ws.range(1, 1, numRows, 1).toString() + ")");
                    ws.formula(numRows + 1, 2, "=SUM(" + ws.range(1, 2, numRows, 2).toString() + ")");
                    ws.formula(numRows + 1, 3, "=SUM(" + ws.range(1, 3, numRows, 3).toString() + ")");
                    ws.formula(numRows + 1, 4, "=AVERAGE(" + ws.range(1, 4, numRows, 4).toString() + ")");
                    ws.style(numRows + 1, 4).format("yyyy-MM-dd HH:mm:ss").set();
                    ws.range(1, 0, numRows, numCols).style().shadeAlternateRows(Color.RED).set();
                });
                cfs[i] = cf;
            }
            try {
                CompletableFuture.allOf(cfs).get();
            } catch (InterruptedException | ExecutionException ex) {
                throw new RuntimeException(ex);
            }
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertEquals(0, xwb.getActiveSheetIndex());
        assertEquals(numWs, xwb.getNumberOfSheets());
        for (int i = 0; i < numWs; ++i) {
            assertEquals("Sheet " + i, xwb.getSheetName(i));
            XSSFSheet xws = xwb.getSheetAt(i);
            assertEquals(numRows + 1, xws.getLastRowNum());
            for (int j = 1; j <= numRows; ++j) {
                assertEquals("String value " + j, xws.getRow(j).getCell(0).getStringCellValue());
            }
        }

    }
}
