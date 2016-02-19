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

import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.Range;
import org.dhatim.fastexcel.Workbook;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
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
    public void colToName() throws Exception {
        assertEquals("AA", Range.colToString(26));
        assertEquals("AAA", Range.colToString(702));
    }
}
