package org.dhatim.fastexcel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class CheckboxWorksheetTest {

    @Test
    public void testCheckboxWorksheet() throws IOException {
        try (ByteArrayOutputStream fos = new ByteArrayOutputStream(); Workbook wb = new Workbook(fos, "FastExcel", "1.0")) {
            Worksheet sheet = wb.newWorksheet("Test");
            sheet.value(1, 1, true);
            sheet.value(1, 2, false);
            sheet.value(2, 1, true);
            sheet.value(2, 2, false);
            sheet.style(2, 1).checkbox(true).set();
            sheet.style(2, 2).checkbox(true).set();
            sheet.close();
        }
    }
}
