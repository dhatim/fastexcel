/*
 * Copyright 2026 Dhatim.
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

import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Comparator;

import static org.assertj.core.api.Assertions.assertThat;

class WorksheetSortTest {

    @ParameterizedTest
    @ValueSource(ints = {0, 1, 3})
    void sortPreservesWorksheetContentAndScopedMetadata(int finishedSheets) throws Exception {
        String[] names = {"AAA", "BBB", "CCC"};
        String[] colors = {"FF0000", "00FF00", "0000FF"};
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (Workbook workbook = new Workbook(output, "Test", "1.0")) {
            Worksheet[] sheets = new Worksheet[names.length];
            for (int i = 0; i < names.length; i++) {
                Worksheet sheet = workbook.newWorksheet(names[i]);
                sheets[i] = sheet;
                sheet.value(0, 0, (i + 1) * 111);
                sheet.setTabColor(colors[i]);
                sheet.comment(0, 0, "Comment for " + names[i]);
                sheet.repeatRows(i);
                sheet.repeatCols(i);
                sheet.setAutoFilter(0, 0, 1, 0);
                Range localRange = sheet.range(i, 0, i, 0);
                localRange.setFolderScope(false);
                sheet.addNamedRange(localRange, "localValue");
                if (i < finishedSheets) {
                    sheet.finish();
                }
            }
            workbook.sortWorksheets(Comparator.comparing(Worksheet::getName).reversed());
            sheets[0].keepInActiveTab();
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            assertThat(workbook.getNumberOfSheets()).isEqualTo(3);
            assertThat(workbook.getSheetName(0)).isEqualTo("CCC");
            assertThat(workbook.getSheetName(1)).isEqualTo("BBB");
            assertThat(workbook.getSheetName(2)).isEqualTo("AAA");
            assertThat(workbook.getActiveSheetIndex()).isEqualTo(2);
            assertThat(workbook.getNames("localValue")).hasSize(3);
            for (int i = 0; i < names.length; i++) {
                XSSFSheet sheet = workbook.getSheet(names[i]);
                assertThat(sheet.getRow(0).getCell(0).getNumericCellValue()).isEqualTo((i + 1) * 111);
                assertThat(sheet.getTabColor().getARGBHex()).endsWith(colors[i]);
                assertThat(sheet.getRow(0).getCell(0).getCellComment().getString().getString())
                        .isEqualTo("Comment for " + names[i]);
                assertThat(sheet.getRepeatingRows().getFirstRow()).isEqualTo(i);
                assertThat(sheet.getRepeatingColumns().getFirstColumn()).isEqualTo(i);
                assertThat(sheet.getCTWorksheet().getAutoFilter().getRef()).isEqualTo("A1:A2");
                int sheetPosition = workbook.getSheetIndex(names[i]);
                XSSFName localName = workbook.getNames("localValue").stream()
                        .filter(name -> name.getSheetIndex() == sheetPosition).findFirst().get();
                assertThat(localName.getRefersToFormula()).startsWith("'" + names[i] + "'!");
            }
        }
    }
}
