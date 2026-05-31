package org.dhatim.fastexcel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.Test;

import java.io.*;

import static org.junit.jupiter.api.Assertions.*;

public class WorkbookProtectionTest {

    private static final File testFile = new File("target/workbookProtectionTest.xlsx");

    private static final String testPassword = "myPassword";

    private static final String testContent = "Hello fastexcel";

    // ── Write helpers ────────────────────────────────────────────────────────

    void fastexcelWriteWithStructureProtection() throws IOException {
        try (FileOutputStream fos = new FileOutputStream(testFile);
             Workbook wb = new Workbook(fos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, testContent);
            wb.protectStructure(testPassword);
        }
    }

    void fastexcelWriteWithoutStructureProtection() throws IOException {
        try (FileOutputStream fos = new FileOutputStream(testFile);
             Workbook wb = new Workbook(fos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, testContent);
        }
    }

    void fastexcelWriteWithNullPassword() throws IOException {
        try (FileOutputStream fos = new FileOutputStream(testFile);
             Workbook wb = new Workbook(fos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, testContent);
            wb.protectStructure(testPassword); // set password
            wb.protectStructure(null);          // then remove it
        }
    }

    // ── Read helpers ─────────────────────────────────────────────────────────

    void poiVerifyStructureIsLocked() throws IOException {
        try (FileInputStream fis = new FileInputStream(testFile);
             XSSFWorkbook poiWb = new XSSFWorkbook(fis)) {
            assertTrue(poiWb.isStructureLocked(),
                "Workbook structure should be locked");
        }
    }

    void poiVerifyStructureIsNotLocked() throws IOException {
        try (FileInputStream fis = new FileInputStream(testFile);
             XSSFWorkbook poiWb = new XSSFWorkbook(fis)) {
            assertFalse(poiWb.isStructureLocked(),
                "Workbook structure should not be locked");
        }
    }

    // ── Cleanup ──────────────────────────────────────────────────────────────

    @AfterAll
    static void cleanup() {
        testFile.delete();
    }

    // ── Tests ────────────────────────────────────────────────────────────────

    @Test
    void fastexcelWrite_poiVerifyStructureLocked() throws Exception {
        fastexcelWriteWithStructureProtection();
        poiVerifyStructureIsLocked();
    }

    @Test
    void fastexcelWrite_poiVerifyStructureNotLocked() throws Exception {
        fastexcelWriteWithoutStructureProtection();
        poiVerifyStructureIsNotLocked();
    }

    @Test
    void fastexcelWrite_nullPassword_poiVerifyStructureNotLocked() throws Exception {
        fastexcelWriteWithNullPassword();
        poiVerifyStructureIsNotLocked();
    }
}