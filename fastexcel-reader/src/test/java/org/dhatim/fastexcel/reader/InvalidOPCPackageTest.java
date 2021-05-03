package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.fail;

public class InvalidOPCPackageTest {
    @Test
    void expectErrors() throws IOException {
        expectError("/invalid/empty.xlsx", "[Content_Types].xml not found");
        expectError("/invalid/only-content-types.xlsx", "/xl/_rels/custom1-workbook.xml.rels not found");
        expectError("/invalid/no-workbook-rels.xlsx", "/xl/_rels/custom1-workbook.xml.rels not found");
        expectError("/invalid/no-workbook-xml.xlsx", "/xl/custom1-workbook.xml not found");
        expectError("/invalid/no-sheet.xlsx", "/xl/worksheets/custom3-sheet1.xml not found");
        expectError("/invalid/missing-sheet-entry.xlsx", "Sheet#0 'Feuil1' is missing an entry in workbook rels (for id: 'rId42')");
    }

    private void expectError(String name, String expected) throws IOException {
        try (ReadableWorkbook wb = new ReadableWorkbook(Resources.open(name))) {
            wb.getFirstSheet().read();
            fail("ExcelReaderException expected");
        } catch (ExcelReaderException ex) {
            assertThat(ex.getMessage()).isEqualTo(expected);
        }
    }

}
