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

import static org.assertj.core.api.Assertions.*;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.Objects;
import java.util.logging.Logger;
import java.util.stream.Stream;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

public class FastExcelReaderTest {

    private static final class RowDates {

        private final int rowNum;
        private final String date1;
        private final String date2;

        public RowDates(int rowNum, String date1, String date2) {
            this.rowNum = rowNum;
            this.date1 = date1;
            this.date2 = date2;
        }

        @Override
        public boolean equals(Object obj) {
            if (this == obj) {
                return true;
            }
            if (obj instanceof RowDates) {
                RowDates other = (RowDates) obj;
                return rowNum == other.rowNum && Objects.equals(date1, other.date1) && Objects.equals(date2, other.date2);
            }
            return false;
        }

        @Override
        public String toString() {
            return "RowDates(" + rowNum + ", " + date1 + ", " + date2 + ")";
        }

    }

    private static final Logger LOGGER = Logger.getLogger(FastExcelReaderTest.class.getName());

    private static final String EXCEL_DATES = "/xlsx/dates.xlsx";

    private static final SimpleDateFormat FORMAT = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm");

    @Test
    public void testDates() throws IOException, OpenXML4JException {
        ArrayList<RowDates> values = new ArrayList<>();

        try (InputStream inputStream = open(EXCEL_DATES); ReadableWorkbook fworkbook = new ReadableWorkbook(inputStream)) {
            try (Stream<Row> stream = fworkbook.getFirstSheet().openStream()) {
                stream.forEach(row -> {
                    values.add(new RowDates(row.getRowNum(), row.getCell(0).asDate().toString(), row.getCell(1).asDate().toString()));
                });
            }
        }

        ArrayList<RowDates> wvalues = new ArrayList<>();
        try (InputStream inputStream = open(EXCEL_DATES); Workbook workbook = WorkbookFactory.create(inputStream)) {
            for (org.apache.poi.ss.usermodel.Row row : workbook.getSheetAt(0)) {
                wvalues.add(new RowDates(row.getRowNum() + 1, toODT(row.getCell(0).getDateCellValue()), toODT(row.getCell(1).getDateCellValue())));
            }
        }

        for (int i = 0; i < values.size(); i++) {
            if (!values.get(i).equals(wvalues.get(i))) {
                assertThat(values.get(i)).isEqualTo(wvalues.get(i));
            }
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {
            "/xlsx/AutoFilter.xlsx",
            "/xlsx/calendar_stress_test.xlsx",
            "/xlsx/cell_style_simple.xlsx",
            "/xlsx/comments_stress_test.xlsx",
            "/xlsx/custom_properties.xlsx",
            "/xlsx/dates.xlsx",
            "/xlsx/defined_names_simple.xlsx",
            "/xlsx/ErrorTypes.xlsx",
            "/xlsx/formula_stress_test.xlsx",
            "/xlsx/formulae_test_simple.xlsx",
            "/xlsx/hyperlink_no_rels.xlsx",
            "/xlsx/hyperlink_stress_test_2011.xlsx",
            "/xlsx/interview.xlsx", "/xlsx/issue.xlsx",
            // "/xlsx/large_strings.xlsx",
            "/xlsx/LONumbers-2010.xlsx",
            "/xlsx/LONumbers-2011.xlsx",
            "/xlsx/LONumbers.xlsx",
            "/xlsx/merge_cells.xlsx",
            "/xlsx/mixed_sheets.xlsx",
            "/xlsx/named_ranges_2011.xlsx",
            "/xlsx/number_format_entities.xlsx",
            "/xlsx/phonetic_text.xlsx",
            "/xlsx/pivot_table_named_range.xlsx",
            "/xlsx/rich_text_stress.xlsx",
            "/xlsx/RkNumber.xlsx",
            "/xlsx/smart_tags_2007.xlsx",
            "/xlsx/sushi.xlsx",
            "/xlsx/text_and_numbers.xlsx",
            "/xlsx/world.xlsx",
            "/xlsx/write.xlsx",
            // "/xlsx/xlsx-stream-d-date-cell.xlsx"
    })
    public void testFile(String file) {
        LOGGER.info("Test " + file);
        try (InputStream inputStream = open(file); InputStream inputStream2 = open(file)) {
            try (ReadableWorkbook excel = new ReadableWorkbook(inputStream); Workbook workbook = WorkbookFactory.create(inputStream2)) {
                Iterator<Sheet> it = excel.getSheets().iterator();
                while (it.hasNext()) {
                    Sheet sheetDef = it.next();

                    org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(sheetDef.getIndex());

                    try (Stream<Row> data = sheetDef.openStream()) {
                        Iterator<Row> rowIt = data.iterator();
                        Iterator<org.apache.poi.ss.usermodel.Row> itr = sheet.iterator();

                        while (rowIt.hasNext()) {
                            Row row = rowIt.next();
                            org.apache.poi.ss.usermodel.Row expected = itr.next();

                            assertThat(row.getPhysicalCellCount()).as("physical cell").isEqualTo(expected.getPhysicalNumberOfCells());
                            assertThat(row.getCellCount()).as("logical cell").isEqualTo(expected.getLastCellNum() == -1 ? 0 : expected.getLastCellNum());

                            for (int i = 0; i < row.getCellCount(); i++) {
                                Cell cell = row.getCell(i);
                                org.apache.poi.ss.usermodel.Cell expCell = expected.getCell(i);

                                assertThat(cell == null).as("cell defined " + i).isEqualTo(expCell == null);
                                if (cell != null) {
                                    String cellAddr = cell.getAddress().toString();
                                    assertThat(toCode(cell.getType())).as("cell type code " + cellAddr).isEqualTo(expCell.getCellTypeEnum().getCode());

                                    if (cell.getType() == CellType.NUMBER) {
                                        BigDecimal n = cell.asNumber();
                                        BigDecimal expN = new BigDecimal(getRawValue(expCell));
                                        assertThat(n).as("Number " + cellAddr).isEqualTo(expN);
                                    } else if (cell.getType() == CellType.STRING) {
                                        String s = cell.asString();
                                        String expS = expCell.getStringCellValue();
                                        assertThat(s).as("String " + cellAddr).isEqualTo(expS);
                                    }
                                }
                            }
                        }
                    } catch (Throwable e) {
                        throw new RuntimeException("On sheet " + sheetDef.getId() + " " + sheetDef.getName(), e);
                    }
                }

            }
        } catch (Throwable e) {
            throw new RuntimeException("On file " + file, e);
        }
    }

    private static int toCode(CellType type) {
        switch (type) {
            case BOOLEAN:
                return 4;
            case EMPTY:
                return 3;
            case ERROR:
                return 5;
            case FORMULA:
                return 2;
            case NUMBER:
                return 0;
            case STRING:
                return 1;
            default:
                throw new IllegalStateException("unknown type " + type);

        }
    }

    private static String toODT(Date date) {
        return FORMAT.format(date);
    }

    private static InputStream open(String name) {
        InputStream result = FastExcelReaderTest.class.getResourceAsStream(name);
        if (result == null) {
            throw new IllegalStateException("Cannot read resource " + name);
        }
        return result;
    }

    private static String getRawValue(org.apache.poi.ss.usermodel.Cell cell) {
        XSSFCell xcell = ((XSSFCell) cell);
        return xcell == null ? null : xcell.getRawValue();
    }

}
