package org.dhatim.fastexcel;

import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.MessageFormat;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertEquals;

public class MemoryUsageE2E {
    private static final File testFile = new File("target/memtest.xlsx");
    public static final int ROWS = 25_000;
    public static final int SHEETS = 3;
    public static final int COLS = 200;
    public static final int FLUSH_EVERY_NR_OR_ROWS = 100;

    @BeforeAll
    static void checkMemoryLimit() {
        assertThat(Runtime.getRuntime().totalMemory() / (1024 * 1024)).isLessThan(51);
    }

    @Test
    void writeAndReadFile() throws IOException {
        try (OutputStream out = new FileOutputStream(testFile)) {
            write(out);
        }

        try (ReadableWorkbook wb = new ReadableWorkbook(testFile)) {
            read(wb);
        }
    }

    private void write(OutputStream out) throws IOException {
        Workbook wb = new Workbook(out, "test", "1.0");
        for (int s = 0; s < SHEETS; s++) {
            Worksheet sheet = wb.newWorksheet("sheet " + s);
            for (int r = 0; r < ROWS; r++) {
                printProgress("writing", s, r);
                for (int c = 0; c < COLS; c++) {
                    sheet.value(r, c, valueFor(r, c));
                }
                if (r % FLUSH_EVERY_NR_OR_ROWS == 0) {
                    sheet.flush();
                }
            }
            sheet.finish();
        }
        wb.finish();
    }

    private void read(ReadableWorkbook wb) {
        wb.getSheets().forEach(sheet -> {
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(r -> {
                    printProgress("reading", sheet.getIndex(), r.getRowNum() - 1);
                    for (int c = 0; c < r.getCellCount(); c++) {
                        assertEquals(
                                valueFor(r.getRowNum() - 1, c),
                                r.getCell(c).asNumber().doubleValue(),
                                1e-5);

                    }
                });
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });
    }

    @AfterAll
    static void cleanup() {
        testFile.delete();
    }

    private static double valueFor(int r, int c) {
        return (double) r * COLS + c;
    }

    private static void printProgress(String prefix, int sheetIndex, int r) {
        int total = ROWS;
        if (r % (total / 100) == 0) {
            String msg = MessageFormat.format("{0} sheet {2}/{3}: {1}%",
                    prefix, 100 * r / total, sheetIndex + 1, SHEETS);
            System.out.println(msg);
        }
    }

}
