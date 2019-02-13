package org.dhatim.fastexcel.reader;

import org.apache.poi.openxml4j.util.ZipSecureFileWorkaround;
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
    private static final int ROWS = 600_001;
    private static final int COLS = 200;
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
                printProgress("writing", r);
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
    public void readFile() throws Exception {
        ZipSecureFileWorkaround.disableZipBombDetection();
        try (ReadableWorkbook wb = new ReadableWorkbook(testFile)) {
            fastexcelReader(wb);
        }
    }

    @Test
    public void readInputStream() throws Exception {
        try (InputStream in = new FileInputStream(testFile);
             ReadableWorkbook wb = new ReadableWorkbook(in)
        ) {
            fastexcelReader(wb);
        }
    }

    private void fastexcelReader(ReadableWorkbook wb) throws IOException {
        Sheet sheet = wb.getFirstSheet();
        try (Stream<Row> rows = sheet.openStream()) {
            rows.forEach(r -> {
                printProgress("reading", r.getRowNum() - 1);
                for (int c = 0; c < r.getCellCount(); c++) {
                    assertEquals(
                            valueFor(r.getRowNum() - 1, c),
                            r.getCell(c).asNumber().doubleValue(),
                            1e-5);

                }
            });
        }
    }

    private static void printProgress(String prefix, int r) {
        if (r % (ROWS / 100) == 0) {
            LOG.info(prefix + ": " + (100 * r / ROWS) + "%");
        }
    }

    private static double valueFor(int r, int c) {
        return (double) r * COLS + c;
    }
}
