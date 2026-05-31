package org.dhatim.fastexcel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ViewPasswordE2ETest {

    @Test
    void testSheetIsHiddenAfterProtectWithViewPassword() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("SecretSheet");
            ws.value(0, 0, "Sensitive Data");
            ws.value(1, 0, "More Sensitive Data");
            ws.protectWithViewPassword("viewPassword");
        }

        byte[] bytes = os.toByteArray();

        // Verify sheet is hidden via Apache POI
        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertTrue(poiWb.isStructureLocked());
        }
    }

    @Test
    void testWorkbookStructureIsLockedAfterProtectWithViewPassword() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("SecretSheet");
            ws.value(0, 0, "Sensitive Data");
            ws.protectWithViewPassword("viewPassword");
        }

        byte[] bytes = os.toByteArray();

        // Verify workbook structure is locked via Apache POI
        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertThat(poiWb.isStructureLocked()).isTrue();
        }
    }

    @Test
    void testDataIsPreservedAfterProtectWithViewPassword() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("SecretSheet");
            ws.value(0, 0, "Sensitive Data");
            ws.value(1, 0, "More Sensitive Data");
            ws.protectWithViewPassword("viewPassword");
        }

        byte[] bytes = os.toByteArray();

        // Verify data is still readable via fastexcel reader
        try (ReadableWorkbook rwb = new ReadableWorkbook(new ByteArrayInputStream(bytes))) {
            try (Stream<Row> rows = rwb.getFirstSheet().openStream()) {
                List<String> values = rows
                        .map(r -> r.getCellAsString(0).orElse(""))
                        .collect(Collectors.toList());
                assertThat(values).containsExactly("Sensitive Data", "More Sensitive Data");
            }
        }
    }

    @Test
    void testOnlyProtectedSheetIsHidden() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet secretSheet = wb.newWorksheet("SecretSheet");
            secretSheet.value(0, 0, "Sensitive Data");
            secretSheet.protectWithViewPassword("viewPassword");

            // Add a second visible sheet
            Worksheet publicSheet = wb.newWorksheet("PublicSheet");
            publicSheet.value(0, 0, "Public Data");
        }

        byte[] bytes = os.toByteArray();

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            // First sheet (SecretSheet) should be hidden
            assertThat(poiWb.isSheetHidden(0)).isTrue();
            // Second sheet (PublicSheet) should be visible
            assertThat(poiWb.isSheetHidden(1)).isFalse();
        }
    }

    @Test
    void testProtectWithViewPasswordAndEditPassword() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("SecretSheet");
            ws.value(0, 0, "Sensitive Data");
            // Protect viewing
            ws.protectWithViewPassword("viewPassword");
            // Also protect editing
            ws.protect("editPassword");
        }

        byte[] bytes = os.toByteArray();

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            // Sheet should be hidden
            assertThat(poiWb.isSheetHidden(0)).isTrue();
            // Workbook structure should be locked
            assertThat(poiWb.isStructureLocked()).isTrue();
            // Sheet should be protected
            assertThat(poiWb.getSheetAt(0).getProtect()).isTrue();
        }
    }
}