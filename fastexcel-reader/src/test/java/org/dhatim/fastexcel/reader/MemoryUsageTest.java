package org.dhatim.fastexcel.reader;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.logging.Logger;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class MemoryUsageTest {
    private static final Logger LOG = Logger.getLogger(MemoryUsageTest.class.getName());
    private static final int ROWS = 500_000;
    private static final int COLS = 100;
    private static File testFile = new File("target/memtest" + ROWS + "x" + COLS + ".xlsx");

    @BeforeAll
    public static void generateBig() throws IOException {
        if (testFile.length() > 0) {
            return;
        }
        LOG.info("Generating " + testFile);
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            SXSSFSheet sheet = wb.createSheet();
            for (int r = 0; r < ROWS; r++) {
                printProgress(r);
                SXSSFRow row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    SXSSFCell cell = row.createCell(c);
                    cell.setCellValue(valueFor(r, c));
                }
            }
            LOG.info("Writing...");
            try (OutputStream out = new FileOutputStream(testFile)) {
                wb.write(out);
            }
        }
        LOG.info("Size: " + testFile.length());
    }

    @Test
    public void read() throws Exception {
        try (ReadableWorkbook wb = new ReadableWorkbook(testFile)) {
            org.dhatim.fastexcel.reader.Sheet sheet = wb.getFirstSheet();
            try (Stream<org.dhatim.fastexcel.reader.Row> rows = sheet.openStream()) {
                rows.forEach(r -> {
                    printProgress(r.getRowNum() - 1);
                    for (int c = 0; c < r.getCellCount(); c++) {
                        assertEquals(
                                valueFor(r.getRowNum() - 1, c),
                                r.getCell(c).asNumber().doubleValue(),
                                1e-5);

                    }
                });
            }
        }
    }

    private static void printProgress(int r) {
        if (r % (ROWS / 100) == 0) {
            LOG.info((100 * r / ROWS) + "%");
        }
    }

    private static double valueFor(int r, int c) {
        return (double) r * COLS + c;
    }
}
