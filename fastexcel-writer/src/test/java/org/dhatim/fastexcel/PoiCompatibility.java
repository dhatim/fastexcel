package org.dhatim.fastexcel;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Date;

import static org.junit.Assert.assertEquals;

public class PoiCompatibility {
    @Test
    public void compatibility() throws IOException {
        byte[] buf = createWithFastexcel();
        try(ByteArrayInputStream in = new ByteArrayInputStream(buf)) {
            validateWithPoi(in);
        }
    }

    private void validateWithPoi(ByteArrayInputStream in) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(in)) {
            assertEquals("number of sheets", 2, workbook.getNumberOfSheets());
            for (int i = 0; i < 2; i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                assertEquals(1, sheet.getLastRowNum());
                XSSFRow row = sheet.getRow(0);
                assertEquals(42.0, row.getCell(0).getNumericCellValue(), 0);
                assertEquals("deadbeef", row.getCell(1).getStringCellValue());
                assertEquals(new Date(1548592289922L), row.getCell(2).getDateCellValue());
            }
        }
    }

    private byte[] createWithFastexcel() throws IOException {
        try(ByteArrayOutputStream buf = new ByteArrayOutputStream()) {
            Workbook wb = new Workbook(buf, "Compat", "1.0");
            for (int i = 0; i < 2; i++) {
                Worksheet ws = wb.newWorksheet("Sheet " + i);
                ws.value(0, 0, 42);
                ws.value(0, 1, "deadbeef");
                ws.value(0, 2, new Date(1548592289922L));
                ws.range(0, 3, 1, 3).style().format("yyyy-mm-dd hh:mm:ss").set();
            }
            wb.finish();
            return buf.toByteArray();
        }
    }

}
