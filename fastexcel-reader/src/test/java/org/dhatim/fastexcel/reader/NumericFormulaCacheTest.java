package org.dhatim.fastexcel.reader;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class NumericFormulaCacheTest {

    @ParameterizedTest
    @ValueSource(strings = {"n", ""})
    void readsNumericFormulaCacheWithoutChangingMetadata(String type) throws IOException {
        byte[] xlsx = workbook(cell(type, "<f>E3</f><v>30</v>"));
        try (ReadableWorkbook workbook = new ReadableWorkbook(new ByteArrayInputStream(xlsx))) {
            Cell cell = workbook.getFirstSheet().read().get(0).getCell(0);
            assertThat(cell.getType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getFormula()).isEqualTo("E3");
            assertThat(cell.getRawValue()).isEqualTo("30");
            assertEquals(new BigDecimal("30"), cell.getValue());
            assertEquals(new BigDecimal("30"), cell.asNumber());
            assertThat(cell.getType()).isEqualTo(CellType.FORMULA);
            assertThrows(ExcelReaderException.class, cell::asString);
            assertThrows(ExcelReaderException.class, cell::asBoolean);
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"n", ""})
    void rowReturnsNumericFormulaCache(String type) throws IOException {
        Row row = read(cell(type, "<f>E3</f><v>30</v>"), false);
        assertThat(row.getCellAsNumber(0)).contains(new BigDecimal("30"));
        assertThat(row.getCell(0).getType()).isEqualTo(CellType.FORMULA);
    }

    @ParameterizedTest
    @ValueSource(strings = {"n", ""})
    void poiRecognizesNumericCacheInTheSameXlsx(String type) throws IOException {
        byte[] xlsx = workbook(cell(type, "<f>E3</f><v>30</v>"));
        try (XSSFWorkbook poi = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            org.apache.poi.ss.usermodel.Cell cell = poi.getSheetAt(0).getRow(2).getCell(0);
            assertThat(cell.getCellType()).isEqualTo(org.apache.poi.ss.usermodel.CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("E3");
            assertThat(cell.getCachedFormulaResultType()).isEqualTo(org.apache.poi.ss.usermodel.CellType.NUMERIC);
            assertThat(cell.getNumericCellValue()).isEqualTo(30);
        }
        Cell cell = read(xlsx, false).getCell(0);
        assertThat(cell.getType()).isEqualTo(CellType.FORMULA);
        assertEquals(new BigDecimal("30"), cell.getValue());
    }

    @ParameterizedTest
    @ValueSource(strings = {"0.000", "-12.3400", "12345678901234567890.1234500", "1.2300E+8"})
    void preservesDecimalPrecisionAndScale(String literal) throws IOException {
        Row row = read(cell("n", "<f>E3</f><v>" + literal + "</v>"), false);
        BigDecimal expected = new BigDecimal(literal);
        assertEquals(expected, row.getCell(0).asNumber());
        assertEquals(expected, row.getCellAsNumber(0).orElseThrow(AssertionError::new));
        assertThat(row.getCell(0).getRawValue()).isEqualTo(literal);
    }

    @ParameterizedTest
    @CsvSource({"str, 30", "b, 1", "e, #DIV/0!"})
    void rejectsNonNumericFormulaCaches(String type, String literal) throws IOException {
        Row row = read(cell(type, "<f>E3</f><v>" + literal + "</v>"), false);
        Cell cell = row.getCell(0);
        assertThat(cell.getType()).isEqualTo(CellType.FORMULA);
        assertThat(cell.getFormula()).isEqualTo("E3");
        assertThat(cell.getRawValue()).isEqualTo(literal);
        assertThat(cell.getValue()).isNotInstanceOf(BigDecimal.class);
        assertThrows(ExcelReaderException.class, cell::asNumber);
        assertThrows(ExcelReaderException.class, () -> row.getCellAsNumber(0));
    }

    @ParameterizedTest
    @ValueSource(strings = {"", "<v/>", "<v></v>"})
    void rejectsMissingOrEmptyFormulaCaches(String cache) throws IOException {
        Row row = read(cell("n", "<f>E3</f>" + cache), false);
        Cell cell = row.getCell(0);
        assertThat(cell.getType()).isEqualTo(CellType.FORMULA);
        assertThat(cell.getFormula()).isEqualTo("E3");
        assertThat(cell.getValue()).isNull();
        assertThrows(ExcelReaderException.class, cell::asNumber);
        assertThrows(ExcelReaderException.class, () -> row.getCellAsNumber(0));
    }

    @Test
    void keepsOrdinaryNumberAndEmptyBehavior() throws IOException {
        Row numeric = read(cell("n", "<v>30.500</v>"), false);
        assertThat(numeric.getCell(0).getType()).isEqualTo(CellType.NUMBER);
        assertEquals(new BigDecimal("30.500"), numeric.getCell(0).asNumber());
        assertEquals(new BigDecimal("30.500"), numeric.getCellAsNumber(0).orElseThrow(AssertionError::new));
        Row empty = read(cell("n", ""), false);
        assertThat(empty.getCell(0).getType()).isEqualTo(CellType.EMPTY);
        assertThat(empty.getCell(0).asNumber()).isNull();
        assertThat(empty.getCellAsNumber(0)).isEmpty();
        assertThat(empty.getCellAsNumber(1)).isEmpty();
    }

    @Test
    void strictReadingRejectsInvalidNumericFormulaCache() {
        assertThrows(ExcelReaderException.class, () -> read(cell("n", "<f>E3</f><v>invalid</v>"), false));
    }

    @Test
    void lenientReadingRejectsInvalidNumericFormulaCache() throws IOException {
        Row row = read(cell("n", "<f>E3</f><v>invalid</v>"), true);
        assertThat(row.getCell(0).getType()).isEqualTo(CellType.FORMULA);
        assertThat(row.getCell(0).getValue()).isNull();
        assertThat(row.getCell(0).getRawValue()).isEqualTo("invalid");
        assertThrows(ExcelReaderException.class, () -> row.getCell(0).asNumber());
        assertThrows(ExcelReaderException.class, () -> row.getCellAsNumber(0));
    }

    @Test
    void lenientReadingDiscardsNumericCacheWhenALaterValueIsInvalid() throws IOException {
        Row row = read(cell("n", "<f>E3</f><v>30</v><v>invalid</v>"), true);
        Cell cell = row.getCell(0);
        assertThat(cell.getType()).isEqualTo(CellType.FORMULA);
        assertThat(cell.getFormula()).isEqualTo("E3");
        assertThat(cell.getRawValue()).isEqualTo("invalid");
        assertThat(cell.getValue()).isNull();
        assertThrows(ExcelReaderException.class, cell::asNumber);
        assertThrows(ExcelReaderException.class, () -> row.getCellAsNumber(0));
    }

    private static Row read(String cell, boolean lenient) throws IOException {
        return read(workbook(cell), lenient);
    }

    private static Row read(byte[] xlsx, boolean lenient) throws IOException {
        try (ReadableWorkbook workbook = new ReadableWorkbook(new ByteArrayInputStream(xlsx), new ReadingOptions(false, lenient))) {
            return workbook.getFirstSheet().read().get(0);
        }
    }

    private static String cell(String type, String content) {
        return "<c r=\"A3\"" + (type.isEmpty() ? "" : " t=\"" + type + "\"") + ">" + content + "</c>";
    }

    private static byte[] workbook(String cell) throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (ZipOutputStream zip = new ZipOutputStream(output)) {
            entry(zip, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                    + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
            entry(zip, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            entry(zip, "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                    + "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            entry(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
            entry(zip, "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"3\">"
                    + cell + "<c r=\"E3\"><v>30</v></c></row></sheetData></worksheet>");
        }
        return output.toByteArray();
    }

    private static void entry(ZipOutputStream zip, String name, String xml) throws IOException {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(xml.getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
    }
}
