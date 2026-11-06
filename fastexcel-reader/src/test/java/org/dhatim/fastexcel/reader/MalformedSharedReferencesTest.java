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
package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MalformedSharedReferencesTest {
    private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    @ParameterizedTest
    @ValueSource(ints = {-1, 1, 4096})
    void rejectsInvalidSharedStringIndex(int index) throws Exception {
        try (ReadableWorkbook workbook = workbook("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>" + index
                + "</v></c></row>", "<si><t>hi</t></si>")) {
            ExcelReaderException exception = assertThrows(ExcelReaderException.class,
                    () -> workbook.getFirstSheet().read());
            assertTrue(exception.getMessage().contains(Integer.toString(index)));
        }
    }

    @Test
    void rejectsIndexIntoEmptySharedStringTable() throws Exception {
        try (ReadableWorkbook workbook = workbook("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>", "")) {
            assertThrows(ExcelReaderException.class, () -> workbook.getFirstSheet().read());
        }
    }

    @Test
    void preservesRichEmptyAndCachedStringsRegardlessOfDeclaredCount() throws Exception {
        String strings = "<si><r><t>rich</t></r><rPh sb=\"0\" eb=\"4\"><t>ignored</t></rPh>"
                + "<r><t> text</t></r></si><si/><si><t>last</t></si>";
        String cells = "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>2</v></c>"
                + "<c r=\"B1\" t=\"s\"><v>0</v></c><c r=\"C1\" t=\"s\"><v>1</v></c></row>";
        try (ReadableWorkbook workbook = workbook(cells, strings)) {
            Row row = workbook.getFirstSheet().read().get(0);
            assertEquals("last", row.getCell(0).asString());
            assertEquals("rich text", row.getCell(1).asString());
            assertEquals("", row.getCell(2).asString());
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"", " si=\"7\""})
    void rejectsSharedFormulaWithoutBase(String index) throws Exception {
        String cells = "<row r=\"1\"><c r=\"A1\"><f t=\"shared\"" + index + "/></c></row>";
        try (ReadableWorkbook workbook = workbook(cells, "")) {
            ExcelReaderException exception = assertThrows(ExcelReaderException.class,
                    () -> workbook.getFirstSheet().read());
            assertTrue(exception.getMessage().contains("Missing base formula"));
            assertTrue(exception.getMessage().contains("A1"));
        }
    }

    @Test
    void preservesValidSharedFormulaExpansion() throws Exception {
        String cells = "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"0\" ref=\"A1:A2\">B1+1</f>"
                + "<v>2</v></c></row><row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>3</v></c></row>";
        try (ReadableWorkbook workbook = workbook(cells, "")) {
            List<Row> rows = workbook.getFirstSheet().read();
            assertEquals("B1+1", rows.get(0).getCell(0).getFormula());
            assertEquals("B2+1", rows.get(1).getCell(0).getFormula());
        }
    }

    private ReadableWorkbook workbook(String rows, String strings) throws IOException {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        try (ZipOutputStream zip = new ZipOutputStream(bytes)) {
            entry(zip, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                    + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
                    + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
            entry(zip, "xl/workbook.xml", "<workbook xmlns=\"" + NS
                    + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                    + "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            entry(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
            // Deliberately inaccurate metadata: actual entries, not uniqueCount, bound reading.
            entry(zip, "xl/sharedStrings.xml", "<sst xmlns=\"" + NS + "\" count=\"1\" uniqueCount=\"1\">" + strings + "</sst>");
            entry(zip, "xl/worksheets/sheet1.xml", "<worksheet xmlns=\"" + NS + "\"><sheetData>" + rows + "</sheetData></worksheet>");
        }
        return new ReadableWorkbook(new ByteArrayInputStream(bytes.toByteArray()));
    }

    private void entry(ZipOutputStream zip, String name, String xml) throws IOException {
        zip.putNextEntry(new ZipEntry(name));
        zip.write(xml.getBytes(StandardCharsets.UTF_8));
        zip.closeEntry();
    }
}
