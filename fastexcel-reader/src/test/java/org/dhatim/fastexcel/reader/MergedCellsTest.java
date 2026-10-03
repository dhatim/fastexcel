package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.dhatim.fastexcel.reader.Resources.open;

class MergedCellsTest {

    private static final List<CellRangeAddress> EXPECTED = Arrays.asList(
            CellRangeAddress.valueOf("A1:B2"),
            CellRangeAddress.valueOf("C1:C2"),
            CellRangeAddress.valueOf("A3:B3"),
            CellRangeAddress.valueOf("D1:D2"),
            CellRangeAddress.valueOf("A4:B4"));

    @Test
    void testMergedCells() throws IOException {
        try (InputStream inputStream = open("/xlsx/merge_cells.xlsx");
             ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
            Sheet sheet = excel.getFirstSheet();
            sheet.read();
            assertThat(sheet.getMergedCells()).isEqualTo(EXPECTED);
        }
    }

    @Test
    void testNoMergedCells() throws IOException {
        try (InputStream inputStream = open("/xlsx/simple.xlsx");
             ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
            Sheet sheet = excel.getFirstSheet();
            sheet.read();
            assertThat(sheet.getMergedCells()).isEmpty();
        }
    }

    @Test
    void testMergedCellsBeforeReadingRows() throws IOException {
        try (InputStream inputStream = open("/xlsx/merge_cells.xlsx");
             ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
            Sheet sheet = excel.getFirstSheet();
            assertThatThrownBy(sheet::getMergedCells).isInstanceOf(IllegalStateException.class);
        }
    }

    @Test
    void testMergedCellsAfterPartialRead() throws IOException {
        try (InputStream inputStream = open("/xlsx/merge_cells.xlsx");
             ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
            Sheet sheet = excel.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                assertThat(rows.findFirst()).isPresent();
            }
            assertThatThrownBy(sheet::getMergedCells).isInstanceOf(IllegalStateException.class);
        }
    }

}
