package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.stream.Stream;

import static org.dhatim.fastexcel.reader.Resources.open;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

public class ExcelFileWithCRLFTest {

    @Test
    public void testFileWithCRLF() throws IOException {
        try (InputStream inputStream = open("/xlsx/withStyleCRLF.xlsx");
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(true, true))) {
            Sheet firstSheet = excel.getFirstSheet();
            try (Stream<Row> rows = firstSheet.openStream()) {
                rows.forEach(r -> {
                    if (r.getRowNum() > 1) {
                        r.forEach(c -> {
                            if (c != null) {
                                assertNotNull(c.getDataFormatString());
                            }
                        });
                    } else {
                        assertEquals("ID", r.getCell(0).getText());
                        assertEquals("Date", r.getCell(1).getText());
                        assertEquals("name", r.getCell(2).getText());
                        assertEquals("valo", r.getCell(3).getText());
                        assertEquals("ratio", r.getCell(4).getText());
                        assertEquals("success", r.getCell(5).getText());
                    }
                });
            }
        }
    }

}
