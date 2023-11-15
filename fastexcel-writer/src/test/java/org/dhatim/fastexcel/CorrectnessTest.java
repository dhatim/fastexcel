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
package org.dhatim.fastexcel;

import org.apache.commons.io.output.NullOutputStream;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.*;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;
import java.util.function.Consumer;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class CorrectnessTest {

    static byte[] writeWorkbook(Consumer<Workbook> consumer) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        Workbook wb = new Workbook(os, "Test", "1.0");
        consumer.accept(wb);
        wb.close();
        return os.toByteArray();
    }

    @Test
    void colToName() {
        assertThat(new Ref() {
        }.colToString(26)).isEqualTo("AA");
        assertThat(new Ref() {
        }.colToString(702)).isEqualTo("AAA");
        assertThat(new Ref() {
        }.colToString(Worksheet.MAX_COLS - 1)).isEqualTo("XFD");
    }

    @Test
    void noWorksheet() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> {
            });
        });
    }

    @Test
    void badVersion() {
        assertThrows(IllegalArgumentException.class, () -> {
            new Workbook(new NullOutputStream(), "Test", "1.0.1");
        });
    }

    @Test
    void singleEmptyWorksheet() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1"));
    }

    @Test
    void worksheetWithNameLongerThan31Chars() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("01234567890123456789012345678901");
            assertThat(ws.getName()).isEqualTo("0123456789012345678901234567890");
        });
    }

    @Test
    void worksheetsWithSameNames() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("01234567890123456789012345678901");
            assertThat(ws.getName()).isEqualTo("0123456789012345678901234567890");
            ws = wb.newWorksheet("0123456789012345678901234567890");
            assertThat(ws.getName()).isEqualTo("01234567890123456789012345678_1");
            ws = wb.newWorksheet("01234567890123456789012345678_1");
            assertThat(ws.getName()).isEqualTo("01234567890123456789012345678_2");
            wb.newWorksheet("abc");
            ws = wb.newWorksheet("abc");
            assertThat(ws.getName()).isEqualTo("abc_1");
        });
    }

    @Test
    void checkMaxRows() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS - 1, 0, "test"));
    }

    @Test
    void checkMaxCols() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS - 1, "test"));
    }

    @Test
    void exceedMaxRows() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS, 0, "test"));
        });
    }

    @Test
    void negativeRow() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(-1, 0, "test"));
        });
    }

    @Test
    void exceedMaxCols() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS, "test"));
        });
    }

    @Test
    void negativeCol() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, -1, "test"));
        });
    }

    @Test
    void invalidRange() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.range(-1, -1, Worksheet.MAX_COLS, Worksheet.MAX_ROWS);
            });
        });
    }

    @Test
    void zoomTooSmall() {
        assertThrows(IllegalArgumentException.class, () -> {
            //if (scale >= 10 && scale <= 400) {
            writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.setZoom(9);
            });
        });
    }

    @Test
    void zoomTooBig() {
        assertThrows(IllegalArgumentException.class, () -> {
            //if (scale >= 10 && scale <= 400) {
            writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.setZoom(401);
            });
        });
    }

    @Test
    void reorderedRange() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            int top = 0;
            int left = 1;
            int bottom = 10;
            int right = 11;
            Range range = ws.range(top, left, bottom, right);
            Range otherRange = ws.range(bottom, right, top, left);
            assertThat(range).isEqualTo(otherRange);
            assertThat(range.getTop()).isEqualTo(top);
            assertThat(range.getLeft()).isEqualTo(left);
            assertThat(range.getBottom()).isEqualTo(bottom);
            assertThat(range.getRight()).isEqualTo(right);
            assertThat(otherRange.getTop()).isEqualTo(top);
            assertThat(otherRange.getLeft()).isEqualTo(left);
            assertThat(otherRange.getBottom()).isEqualTo(bottom);
            assertThat(otherRange.getRight()).isEqualTo(right);
        });
    }

    @Test
    void mergedRanges() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, "One");
            ws.value(0, 1, "Two");
            ws.value(0, 2, "Three");
            ws.value(1, 0, "Merged1");
            ws.value(2, 0, "Merged2");
            ws.range(1, 0, 1, 2).style().merge().set();
            ws.range(2, 0, 2, 2).merge();
            ws.style(1, 0).horizontalAlignment("center").set();
        });
    }

    @Test
    void testForGithubIssue185() throws Exception {
        long start = System.currentTimeMillis();
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            for (int i = 0; i < 10000; i++) {
                if ((i + 1) % 100 == 0) {
                    ws.range(i, 0, i, 19).merge();
                    continue;
                }
                for (int j = 0; j < 20; j++) {
                    ws.value(i, j, "*****");
                }
            }
        });
        long end = System.currentTimeMillis();
        System.out.println("cost:" + (end - start) + "ms");
    }

    @Test
    void testForGithubIssue163() throws Exception {
        // try (FileOutputStream fileOutputStream = new FileOutputStream("D://globalDefaultFontTest.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            wb.setGlobalDefaultFont("Arial", 15.5);
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, "Hello fastexcel");
        });
        // fileOutputStream.write(bytes);
        // }
    }

    @Test
    void testForGithubIssue164() throws Exception {
        // try (FileOutputStream fileOutputStream = new FileOutputStream("D://globalDefaultFontTest.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            wb.setGlobalDefaultFont("Arial", 15.5);
            //General properties
            wb.properties()
                    .setTitle("title property")
                    .setCategory("categrovy property")
                    .setSubject("subject property")
                    .setKeywords("keywords property")
                    .setDescription("description property")
                    .setManager("manager property")
                    .setCompany("company property")
                    .setHyperlinkBase("hyperlinkBase property");
            //Custom properties
            wb.properties()
                    .setTextProperty("Test TextA", "Lucy")
                    .setTextProperty("Test TextB", "Tony")
                    .setDateProperty("Test DateA", Instant.parse("2022-12-22T10:00:00.123456789Z"))
                    .setDateProperty("Test DateB", Instant.parse("1999-09-09T09:09:09Z"))
                    .setNumberProperty("Test NumberA", BigDecimal.valueOf(202222.23364646D))
                    .setNumberProperty("Test NumberB", BigDecimal.valueOf(3.1415926535894D))
                    .setBoolProperty("Test BoolA", true)
                    .setBoolProperty("Test BoolB", false);
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, "Hello fastexcel");
        });
        // fileOutputStream.write(bytes);
        // }
    }

    @Test
    void shouldBeAbleToNullifyCell() throws IOException {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Sheet 1");
            ws.value(0, 0, "One");
            ws.value(1, 0, 42);
            ws.value(2, 0, true);
            ws.value(3, 0, new Date());
            ws.value(4, 0, LocalDate.now());
            ws.value(5, 0, LocalDateTime.now());
            ws.value(6, 0, ZonedDateTime.now());
            for (int r = 0; r <= 6; r++) {
                assertThat(ws.cell(r, 0).getValue()).isNotNull();
            }
            ws.value(0, 0, (Boolean) null);
            ws.value(1, 0, (Number) null);
            ws.value(2, 0, (String) null);
            ws.value(3, 0, (LocalDate) null);
            ws.value(4, 0, (ZonedDateTime) null);
            ws.value(5, 0, (LocalDateTime) null);
            ws.value(6, 0, (LocalDate) null);
            for (int r = 0; r <= 6; r++) {
                assertThat(ws.cell(r, 0).getValue()).isNull();
            }
        });
    }

    @Test
    void testForAllFeatures() throws Exception {
        //The files generated by this test case should always be able to be opened normally with Office software!!
        //try (FileOutputStream fileOutputStream = new FileOutputStream("D://all_features_test.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            //global font
            wb.setGlobalDefaultFont("Arial", 15.5);
            //set property
            wb.properties().setCompany("Test_Company");
            //set custom property
            wb.properties().setTextProperty("Custom_Property_Name", "Custom_Property_Value");
            //create new sheet
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            //hide sheet
            Worksheet invisibleSheet = wb.newWorksheet("Invisible Sheet");
            invisibleSheet.setVisibilityState(VisibilityState.HIDDEN);

            //set tab color
            ws.setTabColor("F381E0");

            //set values for cell
            ws.value(0, 0, "Hello fastexcel");
            ws.inlineString(0, 1, "Hello fastexcel (inline)");
            ws.value(1, 0, 1024.2048);
            ws.value(2, 0, true);
            ws.value(3, 0, new Date());
            ws.value(4, 0, LocalDateTime.now());
            ws.value(5, 0, LocalDate.now());
            ws.value(6, 0, ZonedDateTime.now());
            //set hyperlink for cell
            ws.hyperlink(7, 0, new HyperLink("https://github.com/dhatim/fastexcel", "Test_Hyperlink_For_Cell"));
            //set comment for cell
            ws.comment(8, 0, "Test_Comment");
            //set header and footer
            ws.header("Test_Header", Position.CENTER);
            ws.footer("Test_Footer", Position.LEFT);
            //set hide row
            ws.value(9, 0, "You can't see this row");
            ws.hideRow(9);
            //set column width
            ws.width(1, 20);
            //set zoom
            ws.setZoom(120);
            //set formula
            ws.value(10, 0, 44444);
            ws.value(10, 1, 55555);
            ws.formula(10, 2, "=SUM(A11,B11)");
            //set style for cell
            ws.value(11, 0, "Test_Cell_Style");
            ws.style(11, 0)
                    .borderStyle(BorderSide.DIAGONAL, BorderStyle.MEDIUM)
                    .fontSize(20)
                    .fontColor(Color.RED)
                    .italic()
                    .bold()
                    .fillColor(Color.YELLOW)
                    .fontName("Cooper Black")
                    .borderColor(Color.SEA_BLUE)
                    .underlined()
                    .rotation(90)
                    .set();
            //merge cells
            ws.range(12, 0, 12, 3).merge();
            //set hyperlink for range
            ws.range(13, 0, 13, 3).setHyperlink(new HyperLink("https://github.com/dhatim/fastexcel", "Test_Hyperlink_For_Range"));
            //set name for range
            ws.range(14, 0, 14, 3).setName("Test_Set_Name");
            //set style for range
            ws.value(15, 0, "Test_Range_Style");
            ws.range(15, 0, 19, 3).style()
                    .borderStyle(BorderSide.DIAGONAL, BorderStyle.MEDIUM)
                    .fontSize(20)
                    .fontColor(Color.RED)
                    .italic()
                    .bold()
                    .fillColor(Color.YELLOW)
                    .fontName("Cooper Black")
                    .borderColor(Color.SEA_BLUE)
                    .underlined()
                    .shadeAlternateRows(Color.SEA_BLUE)
                    .shadeRows(Color.RED, 1)
                    .set();
            //protect the sheet
            ws.protect("1234");
            //autoFilter
            ws.value(21, 0, "A");
            ws.value(21, 1, "A");
            ws.value(21, 2, "A");
            ws.value(22, 0, "B");
            ws.value(22, 1, "B");
            ws.value(22, 2, "B");
            ws.setAutoFilter(20, 0, 22, 2);

            //validation
            ws.value(23, 0, "ABAB");
            ws.value(23, 1, "CDCD");
            ws.value(24, 0, "EFEF");
            ws.value(24, 1, "GHGH");
            ws.range(25, 0, 25, 1).validateWithList(ws.range(23, 0, 24, 1));

            //table
            ws.range(0, 0, 5, 2).createTable();

            //group
            ws.groupRows(3,4);
            ws.groupRows(2,5);

            ws.groupCols(6,7);
            ws.groupCols(4,8);
        });

        //fileOutputStream.write(bytes);
        //}
    }

    @Test
    void testForGithubIssue72() throws Exception {
        //try (FileOutputStream fileOutputStream = new FileOutputStream("D://fast_hyperlink.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            wb.setGlobalDefaultFont("Arial", 15.5);
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.hyperlink(0, 0, new HyperLink("https://github.com/dhatim/fastexcel", "Baidu"));
            ws.range(1, 0, 1, 1).setHyperlink(new HyperLink("./dev_soft/test.pdf", "dev_folder"));
        });
        //fileOutputStream.write(bytes);
        //}
    }

    @Test
    void testForGithubIssue182() throws Exception {
        //try (FileOutputStream fileOutputStream = new FileOutputStream("D://fast_table.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.range(0, 0, 5, 2).createTable()
                    .setDisplayName("TableDisplayName")
                    .setName("TableName")
                    .styleInfo()
                    .setStyleName("TableStyleMedium1")
                    .setShowLastColumn(true);
        });
        //fileOutputStream.write(bytes);
        //}
    }

    @Test
    void testForGithubIssue254() throws Exception {
        assertThrows(IllegalArgumentException.class, () -> {
            //try (FileOutputStream fileOutputStream = new FileOutputStream("D://merge_conflicted.xlsx")) {
            byte[] bytes = writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.range(0, 0, 2, 2).merge();
                ws.range(1, 1, 3, 3).merge();
            });
            //fileOutputStream.write(bytes);
            //}
        });
    }


    @Test
    void testForTableConflict() throws Exception {
        assertThrows(IllegalArgumentException.class, () -> {
            //try (FileOutputStream fileOutputStream = new FileOutputStream("D://table_conflicted.xlsx")) {
            byte[] bytes = writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.range(0, 0, 2, 2).createTable();
                ws.range(1, 1, 3, 3).createTable();
            });
            //fileOutputStream.write(bytes);
            //}
        });
    }

    @Test
    void testForOffBy1ErrorFor1900_localDateTime() {
        LocalDateTime ldt1 = LocalDateTime.of(1900, Month.JANUARY, 1, 0, 0);
        LocalDateTime ldt2 = LocalDateTime.of(1901, Month.JANUARY, 1, 0, 0);
        LocalDateTime ldt3 = LocalDateTime.of(2000, Month.JANUARY, 1, 0, 0);
        LocalDateTime ldt4 = LocalDateTime.of(2023, Month.JANUARY, 1, 0, 0);
        LocalDateTime ldt5 = LocalDateTime.of(1960, Month.JANUARY, 1, 0, 0);
        assertThat(TimestampUtil.convertDate(ldt1)).isEqualTo(1.0);
        assertThat(TimestampUtil.convertDate(ldt2)).isEqualTo(367.0);
        assertThat(TimestampUtil.convertDate(ldt3)).isEqualTo(36526.0);
        assertThat(TimestampUtil.convertDate(ldt4)).isEqualTo(44927.0);
        assertThat(TimestampUtil.convertDate(ldt5)).isEqualTo(21916.0);
    }

    @Test
    void testForOffBy1ErrorFor1900_localDate() {
        LocalDate ldt1 = LocalDate.of(1900, Month.JANUARY, 1);
        LocalDate ldt2 = LocalDate.of(1901, Month.JANUARY, 1);
        LocalDate ldt3 = LocalDate.of(2000, Month.JANUARY, 1);
        LocalDate ldt4 = LocalDate.of(2023, Month.JANUARY, 1);
        LocalDate ldt5 = LocalDate.of(1960, Month.JANUARY, 1);
        assertThat(TimestampUtil.convertDate(ldt1)).isEqualTo(1.0);
        assertThat(TimestampUtil.convertDate(ldt2)).isEqualTo(367.0);
        assertThat(TimestampUtil.convertDate(ldt3)).isEqualTo(36526.0);
        assertThat(TimestampUtil.convertDate(ldt4)).isEqualTo(44927.0);
        assertThat(TimestampUtil.convertDate(ldt5)).isEqualTo(21916.0);
    }

    @Test
    void testForOffBy1ErrorFor1900_utilDate() {
        Date d1 = getCalendarDate(1900, 1, 1);
        Date d2 = getCalendarDate(1901, 1, 1);
        Date d3 = getCalendarDate(2000, 1, 1);
        Date d4 = getCalendarDate(2023, 1, 1);
        Date d5 = getCalendarDate(1960, 1, 1);
        System.out.println(d1);
        assertThat(TimestampUtil.convertDate(d1)).isEqualTo(1.0);
        assertThat(TimestampUtil.convertDate(d2)).isEqualTo(367.0);
        assertThat(TimestampUtil.convertDate(d3)).isEqualTo(36526.0);
        assertThat(TimestampUtil.convertDate(d4)).isEqualTo(44927.0);
        assertThat(TimestampUtil.convertDate(d5)).isEqualTo(21916.0);
    }

    private static Date getCalendarDate(int year, int month, int day) {
        Calendar cal = Calendar.getInstance();
        cal.setTimeZone(TimeZone.getDefault());
        cal.set(Calendar.YEAR, year);
        cal.set(Calendar.MONTH, month - 1);
        cal.set(Calendar.DAY_OF_MONTH, day);
        cal.set(Calendar.HOUR_OF_DAY, 0);
        cal.set(Calendar.MINUTE, 0);
        cal.set(Calendar.SECOND, 0);
        cal.set(Calendar.MILLISECOND, 0);
        Date time = cal.getTime();
        return time;
    }


    @Test
    void testForIssue261() throws Exception {
        //try (FileOutputStream fileOutputStream = new FileOutputStream("D://hide_test.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(1, 1, "hidden cell");
            ws.hideRow(1);
            ws.hideColumn(1);

            ws.hideRow(3);
            ws.hideColumn(3);
        });
        //fileOutputStream.write(bytes);
        //}
    }

    @Test
    void testForIssue259() throws Exception {
        //try (FileOutputStream fileOutputStream = new FileOutputStream("D://group_cols_test.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(1,1,"tt");
            ws.value(1,2,"tt");
            ws.value(1,3,"tt");
            ws.value(1,4,"tt");
            ws.value(1,5,"tt");

            ws.value(1,0,"cc");
            ws.value(2,0,"cc");
            ws.value(3,0,"cc");
            ws.value(4,0,"cc");
            ws.value(5,0,"cc");

            ws.groupCols(2,3);
            ws.groupCols(1,5);

            ws.groupRows(2,3);
            ws.groupRows(1,5);

        });
        //fileOutputStream.write(bytes);
        //}
    }

    @Test
    void testForTimeZoneWhenUsingZonedDateTime() {
        final Instant instant = Instant.ofEpochSecond(0);
        final ZonedDateTime utc = ZonedDateTime.ofInstant(instant, ZoneId.of("Z"));
        final ZonedDateTime two = ZonedDateTime.ofInstant(instant, ZoneId.of("+02:00"));

        assertEquals(25569.0, TimestampUtil.convertZonedDateTime(utc));
        assertEquals(25569.083, TimestampUtil.convertZonedDateTime(two), 0.001);
    }

    @Test
    void testForIssue285() throws Exception {
        // try (FileOutputStream fileOutputStream = new FileOutputStream("D://global_font_test.xlsx")) {
        byte[] bytes = writeWorkbook(wb -> {
            Worksheet sheet = wb.newWorksheet("test sheet");
            wb.setGlobalDefaultFont("Arial", 15.5);

            sheet.value(0, 0, "should be arial"); // but it is Calibri
            sheet.style(0, 0).fontColor(Color.RED).set();

            sheet.value(2, 0, "no additional style"); // Arial 15.5 as global default

            sheet.value(3, 0, "manual arial with style");
            sheet.style(3, 0).fontName("Arial").fontColor(Color.GREEN).set(); // Arial 11

        });
        // fileOutputStream.write(bytes);
        // }
    }

    @Test
    public void testDiagonalProperties() throws Exception {
        writeWorkbook(wb -> {
            Worksheet worksheet = wb.newWorksheet("Worksheet 1");
            worksheet.style(0, 0)
                    .borderColor(BorderSide.RIGHT, Color.BLACK)
                    .borderStyle(BorderSide.RIGHT, BorderStyle.THIN)
                    .borderColor(BorderSide.DIAGONAL, Color.BLACK)
                    .borderStyle(BorderSide.DIAGONAL, BorderStyle.THIN)
                    .diagonalProperty(DiagonalProperty.DIAGONAL_UP)
                    .set();
            worksheet.style(1, 0)
                    .borderColor(BorderSide.TOP, Color.BLACK)
                    .borderStyle(BorderSide.TOP, BorderStyle.THIN)
                    .borderColor(BorderSide.DIAGONAL, Color.BLACK)
                    .borderStyle(BorderSide.DIAGONAL, BorderStyle.MEDIUM)
                    .diagonalProperty(DiagonalProperty.DIAGONAL_DOWN)
                    .set();
            worksheet.style(2, 0)
                    .borderColor(BorderSide.TOP, Color.BLACK)
                    .borderStyle(BorderSide.TOP, BorderStyle.THIN)
                    .borderColor(BorderSide.DIAGONAL, Color.BLACK)
                    .borderStyle(BorderSide.DIAGONAL, BorderStyle.MEDIUM)
                    .diagonalProperty(DiagonalProperty.DIAGONAL_DOWN)
                    .diagonalProperty(DiagonalProperty.DIAGONAL_UP)
                    .diagonalProperty(DiagonalProperty.DIAGONAL_UP)
                    .set();
        });
    }

    @Test
    void testForIndent() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, "One");
            ws.value(0, 1, "Two");
            ws.value(0, 2, "Three");
            ws.value(1, 2, "Test");
            ws.style(0, 0).indent(1).set();
            ws.range(0, 1, 0, 2).style().indent(5).set();
            ws.style(1, 2).indent(16).set();
            ws.style(2,0).indent(0).set();
        });
    }

    @Test
    void testCustomValidation() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Sheet 1");

            ws.range(0, 0, 10, 0).validateWithFormula("IsNumber(A1)")
                    .allowBlank(false)
                    .errorTitle("Error")
                    .error("Wrong value")
                    .showErrorMessage(true)
                    .errorStyle(DataValidationErrorStyle.STOP);

            //Is number
            ws.range(0, 1, 10, 1).validateWithFormula("mod(B1,1)=0")
                    .allowBlank(false)
                    .errorTitle("Error")
                    .error("Wrong value")
                    .showErrorMessage(true)
                    .errorStyle(DataValidationErrorStyle.STOP);

            //Is date
            ws.range(0, 2, 10, 2).validateWithFormula("ISNUMBER(DATEVALUE(TEXT(C1, \"dd/mm/yyyy\")))")
                    .allowBlank(false)
                    .errorTitle("Error")
                    .error("Wrong value")
                    .showErrorMessage(true)
                    .errorStyle(DataValidationErrorStyle.STOP);

            //Is bool
            ws.range(0, 3, 10, 3).validateWithFormula("isLogical(D1)")
                    .allowBlank(false)
                    .errorTitle("Error")
                    .error("Wrong value")
                    .showErrorMessage(true)
                    .errorStyle(DataValidationErrorStyle.STOP);
        });
    }

    @Test
    void testInternalHyperlinks() throws Exception{
        writeWorkbook(wb -> {
            Worksheet worksheet1 = wb.newWorksheet("Sheet1");
            Worksheet worksheet2 = wb.newWorksheet("Sheet2");

            worksheet1.hyperlink(1, 1, HyperLink.internal("Sheet2!A1", "HyperLink"));
            worksheet1.hyperlink(7, 0, HyperLink.external("https://github.com/dhatim/fastexcel", "Test_Hyperlink_For_Cell"));
        });
    }
}
