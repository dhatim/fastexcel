package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;

import static org.assertj.core.api.Assertions.assertThat;
import static org.dhatim.fastexcel.reader.Resources.file;
import static org.dhatim.fastexcel.reader.Resources.open;

public class WithFormatTest {

    private static final String WITH_STYLE_XLSX = "/xlsx/withStyle.xlsx";

    @Test
    void testWithStyleWithCellFormatWorkbookFromInputStream() throws IOException {
        try (InputStream inputStream = open(WITH_STYLE_XLSX);
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(true, false))) {
            assertWithStyleWithCellFormat(excel);
        }
    }

    @Test
    void testWithStyleWithoutCellFormatWorkbookFromInputStream() throws IOException {
        try (InputStream inputStream = open(WITH_STYLE_XLSX);
             ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(false, false))) {
            assertWithStyleWithoutCellFormat(excel);
        }
    }

    @Test
    void testWithStyleWithCellFormatWorkbookFromFile() throws IOException {
        try (ReadableWorkbook excel = new ReadableWorkbook(file(WITH_STYLE_XLSX),
                new ReadingOptions(true, false))) {
            assertWithStyleWithCellFormat(excel);
        }
    }

    @Test
    void testWithStyleWithoutCellFormatWorkbookFromFile() throws IOException {
        try (ReadableWorkbook excel = new ReadableWorkbook(file(WITH_STYLE_XLSX),
                new ReadingOptions(false, false))) {
            assertWithStyleWithoutCellFormat(excel);
        }
    }

    private void assertWithStyleWithCellFormat(ReadableWorkbook excel) throws IOException {
        Optional<Sheet> sheet = excel.getActiveSheet();
        assertThat(sheet).isPresent();
        Row[] rows = sheet.get().openStream().toArray(Row[]::new);
        assertThat(rows).hasSizeGreaterThanOrEqualTo(4);
        assertThat(rows[1].getCell(0)).extracting(Cell::getDataFormatId).isEqualTo(164);
        assertThat(rows[1].getCell(0)).extracting(Cell::getDataFormatString).isEqualTo("General");
        assertThat(rows[2].getCell(1)).extracting(Cell::getDataFormatId).isEqualTo(166);
        assertThat(rows[2].getCell(1)).extracting(Cell::getDataFormatString).isEqualTo("DD\\-MM\\-YYYY");
        assertThat(rows[3].getCell(2)).extracting(Cell::getDataFormatId).isEqualTo(165);
        assertThat(rows[3].getCell(2)).extracting(Cell::getDataFormatString).isEqualTo("D\". \"MMMM\\ YYYY");
    }

    private void assertWithStyleWithoutCellFormat(ReadableWorkbook excel) throws IOException {
        Optional<Sheet> sheet = excel.getActiveSheet();
        assertThat(sheet).isPresent();
        sheet.get().openStream().forEach(row -> row.stream().forEach(cell -> {
            assertThat(cell).extracting(Cell::getDataFormatId).isNull();
            assertThat(cell).extracting(Cell::getDataFormatString).isNull();
        }));
    }
}
