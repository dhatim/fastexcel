/*
 * Copyright 2018 Dhatim.
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
package org.dhatim.fastexcel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

class WriterTest {

    @Test
    void testEscaping() throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        Writer w = new Writer(baos);
        w.append("not escaped");
        w.appendEscaped(" but <this will be escaped \ud83d\ude01>");
        w.flush();
        String s = baos.toString("UTF-8");
        assertThat(s).isEqualTo("not escaped but &lt;this will be escaped &#x1f601;&gt;");
    }

    @Test
    void testEscapingInvalidCharacters() throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        Writer w = new Writer(baos);
        w.appendEscaped("some characters are ignored: \b or \u0001");
        w.flush();
        String s = baos.toString("UTF-8");
        assertThat(s).isEqualTo("some characters are ignored:  or ");
    }
    @Test
    void protectStructureLocksWorkbookStructure() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, "Hello");

            wb.protectStructure("myPassword");
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(
                new ByteArrayInputStream(os.toByteArray()))) {
            assertThat(poiWb.isStructureLocked()).isTrue();
         }
    }
    @Test
    void protectStructureWithNullPasswordDoesNotLockWorkbookStructure() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, "Hello");

            wb.protectStructure(null);
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(
                new ByteArrayInputStream(os.toByteArray()))) {
            assertThat(poiWb.isStructureLocked()).isFalse();
        }
    }
    @Test
    void protectWithViewPasswordOnlyHidesTargetSheet() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(os, "Test", "1.0")) {
            Worksheet secretSheet = wb.newWorksheet("SecretSheet");
            secretSheet.value(0, 0, "Sensitive Data");
            secretSheet.protectWithViewPassword("viewPassword");

            Worksheet publicSheet = wb.newWorksheet("PublicSheet");
            publicSheet.value(0, 0, "Public Data");
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(
                new ByteArrayInputStream(os.toByteArray()))) {
            assertThat(poiWb.isSheetHidden(0)).isTrue();
            assertThat(poiWb.isSheetHidden(1)).isFalse();
        }
    }

}
