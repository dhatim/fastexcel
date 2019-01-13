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

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Date;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.function.Consumer;

import org.apache.commons.io.output.NullOutputStream;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;

public class Correctness {

    private byte[] writeWorkbook(Consumer<Workbook> consumer) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        Workbook wb = new Workbook(os, "Test", "1.0");
        consumer.accept(wb);
        wb.finish();
        return os.toByteArray();
    }

    @Test
    public void colToName() throws Exception {
        assertThat(Range.colToString(26)).isEqualTo("AA");
        assertThat(Range.colToString(702)).isEqualTo("AAA");
        assertThat(Range.colToString(Worksheet.MAX_COLS - 1)).isEqualTo("XFD");
    }

    @Test(expected = IllegalArgumentException.class)
    public void noWorksheet() throws Exception {
        writeWorkbook(wb -> {
        });
    }

    @Test(expected = IllegalArgumentException.class)
    public void badVersion() throws Exception {
        Workbook dummy = new Workbook(new NullOutputStream(), "Test", "1.0.1");
    }

    @Test
    public void singleEmptyWorksheet() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1"));
    }

    @Test
    public void worksheetWithNameLongerThan31Chars() throws Exception {
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("01234567890123456789012345678901");
            assertThat(ws.getName()).isEqualTo("0123456789012345678901234567890");
        });
    }

    @Test
    public void worksheetsWithSameNames() throws Exception {
        byte[] data = writeWorkbook(wb -> {
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
    public void checkMaxRows() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS - 1, 0, "test"));
    }

    @Test
    public void checkMaxCols() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS - 1, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void exceedMaxRows() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS, 0, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void negativeRow() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(-1, 0, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void exceedMaxCols() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void negativeCol() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, -1, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void notSupportedTypeCell() throws Exception {
        byte[] data = writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, 0, new Object()));
    }

    @Test(expected = IllegalArgumentException.class)
    public void invalidRange() throws Exception {
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.range(-1, -1, Worksheet.MAX_COLS, Worksheet.MAX_ROWS);
        });
    }

    @Test
    public void reorderedRange() throws Exception {
        byte[] data = writeWorkbook(wb -> {
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
    public void mergedRanges() throws Exception {
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, "One");
            ws.value(0, 1, "Two");
            ws.value(0, 2, "Three");
            ws.value(1, 0, "Merged");
            ws.range(1, 0, 1, 2).style().merge().set();
            ws.range(1, 0, 1, 2).merge();
            ws.style(1, 0).horizontalAlignment("center").set();
        });
    }

    @Test
    public void singleWorksheet() throws Exception {
        String sheetName = "Worksheet 1";
        String stringValue = "Sample text with chars to escape : < > & \\ \" ' ~ é è à ç ù µ £ €";
        Date dateValue = new Date();
        // We truncate to milliseconds because JDK 9 gives microseconds, whereas Excel format supports only milliseconds
        LocalDateTime localDateTimeValue = LocalDateTime.now().truncatedTo(ChronoUnit.MILLIS);
        ZoneId timezone = ZoneId.of("Australia/Sydney");
        ZonedDateTime zonedDateValue = ZonedDateTime.ofInstant(dateValue.toInstant().truncatedTo(ChronoUnit.MILLIS), timezone);
        double doubleValue = 1.234;
        int intValue = 2_016;
        long longValue = 2_016_000_000_000L;
        BigDecimal bigDecimalValue = BigDecimal.TEN;
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet(sheetName);
            ws.width(0, 2);
            int i = 1;
            ws.hideRow(i);
            ws.value(i, i++, stringValue);
            ws.value(i, i++, dateValue);
            ws.value(i, i++, localDateTimeValue);
            ws.value(i, i++, zonedDateValue);
            ws.value(i, i++, doubleValue);
            ws.value(i, i++, intValue);
            ws.value(i, i++, longValue);
            ws.value(i, i++, bigDecimalValue);
            try {
                ws.finish();
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertThat(xwb.getActiveSheetIndex()).isEqualTo(0);
        assertThat(xwb.getNumberOfSheets()).isEqualTo(1);
        XSSFSheet xws = xwb.getSheet(sheetName);
        @SuppressWarnings("unchecked")
        Comparable<XSSFRow> row = (Comparable) xws.getRow(0);
        assertThat(row).isNull();
        int i = 1;
        // poi column width is in 1/256 characters
        assertThat(xws.getColumnWidth(0) / 256).isEqualTo(2);
        assertThat(xws.getRow(i).getZeroHeight()).isTrue();
        assertThat(xws.getRow(i).getCell(i++).getStringCellValue()).isEqualTo(stringValue);
        assertThat(xws.getRow(i).getCell(i++).getDateCellValue()).isEqualTo(dateValue);
        // Check zoned timestamps have the same textual representation as the Dates extracted from the workbook
        // (Excel date serial numbers do not carry timezone information)
        assertThat(DateTimeFormatter.ISO_LOCAL_DATE_TIME.format(ZonedDateTime.ofInstant(xws.getRow(i).getCell(i++).getDateCellValue().toInstant(), ZoneId.systemDefault()))).isEqualTo(DateTimeFormatter.ISO_LOCAL_DATE_TIME.format(localDateTimeValue));
        assertThat(DateTimeFormatter.ISO_LOCAL_DATE_TIME.format(ZonedDateTime.ofInstant(xws.getRow(i).getCell(i++).getDateCellValue().toInstant(), ZoneId.systemDefault()))).isEqualTo(DateTimeFormatter.ISO_LOCAL_DATE_TIME.format(zonedDateValue));
        assertThat(xws.getRow(i).getCell(i++).getNumericCellValue()).isEqualTo(doubleValue);
        assertThat(xws.getRow(i).getCell(i++).getNumericCellValue()).isEqualTo(intValue);
        assertThat(xws.getRow(i).getCell(i++).getNumericCellValue()).isEqualTo(longValue);
        assertThat(new BigDecimal(xws.getRow(i).getCell(i++).getRawValue())).isEqualTo(bigDecimalValue);
    }

    @Test
    public void multipleWorksheets() throws Exception {
        int numWs = 10;
        int numRows = 5000;
        int numCols = 6;
        byte[] data = writeWorkbook(wb -> {
            @SuppressWarnings("unchecked")
            CompletableFuture<Void>[] cfs = new CompletableFuture[numWs];
            for (int i = 0; i < cfs.length; ++i) {
                Worksheet ws = wb.newWorksheet("Sheet " + i);
                CompletableFuture<Void> cf = CompletableFuture.runAsync(() -> {
                    for (int j = 0; j < numCols; ++j) {
                        ws.value(0, j, "Column " + j);
                        ws.style(0, j).bold().fontSize(12).fillColor(Color.GRAY2).set();
                        for (int k = 1; k <= numRows; ++k) {
                            switch (j) {
                                case 0:
                                    ws.value(k, j, "String value " + k);
                                    break;
                                case 1:
                                    ws.value(k, j, 2);
                                    break;
                                case 2:
                                    ws.value(k, j, 3L);
                                    break;
                                case 3:
                                    ws.value(k, j, 0.123);
                                    break;
                                case 4:
                                    ws.value(k, j, new Date());
                                    ws.style(k, j).format("yyyy-MM-dd HH:mm:ss").set();
                                    break;
                                case 5:
                                    ws.value(k, j, LocalDate.now());
                                    ws.style(k, j).format("yyyy-MM-dd").set();
                                    break;
                                default:
                                    throw new IllegalArgumentException();
                            }
                        }
                    }
                    ws.formula(numRows + 1, 1, "=SUM(" + ws.range(1, 1, numRows, 1).toString() + ")");
                    ws.formula(numRows + 1, 2, "=SUM(" + ws.range(1, 2, numRows, 2).toString() + ")");
                    ws.formula(numRows + 1, 3, "=SUM(" + ws.range(1, 3, numRows, 3).toString() + ")");
                    ws.formula(numRows + 1, 4, "=AVERAGE(" + ws.range(1, 4, numRows, 4).toString() + ")");
                    ws.style(numRows + 1, 4).format("yyyy-MM-dd HH:mm:ss").set();
                    ws.formula(numRows + 1, 5, "=AVERAGE(" + ws.range(1, 5, numRows, 5).toString() + ")");
                    ws.style(numRows + 1, 5).format("yyyy-MM-dd").bold().italic().fontColor(Color.RED).fontName("Garamond").fontSize(new BigDecimal("14.5")).horizontalAlignment("center").verticalAlignment("top").wrapText(true).set();
                    ws.range(1, 0, numRows, numCols - 1).style().borderColor(Color.RED).borderStyle("thick").shadeAlternateRows(Color.RED).set();
                });
                cfs[i] = cf;
            }
            try {
                CompletableFuture.allOf(cfs).get();
            } catch (InterruptedException | ExecutionException ex) {
                throw new RuntimeException(ex);
            }
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertThat(xwb.getActiveSheetIndex()).isEqualTo(0);
        assertThat(xwb.getNumberOfSheets()).isEqualTo(numWs);
        for (int i = 0; i < numWs; ++i) {
            assertThat(xwb.getSheetName(i)).isEqualTo("Sheet " + i);
            XSSFSheet xws = xwb.getSheetAt(i);
            assertThat(xws.getLastRowNum()).isEqualTo(numRows + 1);
            for (int j = 1; j <= numRows; ++j) {
                assertThat(xws.getRow(j).getCell(0).getStringCellValue()).isEqualTo("String value " + j);
            }
        }
    }


    @Test
    public void sortWorksheets() throws Exception {
        int numWs = 3;
        byte[] data = writeWorkbook(wb -> {
            for (int i = 0; i < numWs; ++i) {
                wb.newWorksheet("Sheet " + i);
            }
            // sort sheet in reverse order
            wb.sortWorksheets((s1, s2) -> -(s1.getName().compareTo(s2.getName())));
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertThat(xwb.getActiveSheetIndex()).isEqualTo(0);
        assertThat(xwb.getNumberOfSheets()).isEqualTo(numWs);
        int nameI = numWs - 1;
        for (int i = 0; i > numWs; ++i) {
            assertThat(xwb.getSheetName(i)).isEqualTo("Sheet " + nameI);
            --nameI;
        }
    }

    @Test
    public void borders() throws Exception {
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.style(1, 1).borderStyle("thin").set();
            ws.style(1, 2).borderStyle("thick").borderColor(Color.RED).set();
            ws.style(1, 3).borderStyle(BorderSide.BOTTOM, "thick").borderColor(BorderSide.BOTTOM, Color.RED).set();
            ws.style(1, 4).borderStyle(BorderSide.TOP, "thin").set();
            ws.style(1, 5).borderStyle(BorderSide.LEFT, "thin").borderStyle(BorderSide.BOTTOM, "thick").set();
            ws.style(1, 6).borderStyle(BorderSide.RIGHT, "thin").set();
            for (int col = 1; col < 10; ++col) {
                ws.style(2, col).borderStyle(BorderSide.LEFT, "thin").borderStyle(BorderSide.BOTTOM, "thin").set();
            }
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertThat(xwb.getActiveSheetIndex()).isEqualTo(0);
        assertThat(xwb.getNumberOfSheets()).isEqualTo(1);
        XSSFSheet xws = xwb.getSheetAt(0);
        assertThat(xws.getLastRowNum()).isEqualTo(2);

        assertThat(xws.getRow(1).getCell(1).getCellStyle().getTopBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(1).getCellStyle().getLeftBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(1).getCellStyle().getBottomBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(1).getCellStyle().getRightBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(1).getCellStyle().getBorderTopEnum()).isEqualTo(BorderStyle.THIN);
        assertThat(xws.getRow(1).getCell(1).getCellStyle().getBorderLeftEnum()).isEqualTo(BorderStyle.THIN);
        assertThat(xws.getRow(1).getCell(1).getCellStyle().getBorderBottomEnum()).isEqualTo(BorderStyle.THIN);
        assertThat(xws.getRow(1).getCell(1).getCellStyle().getBorderRightEnum()).isEqualTo(BorderStyle.THIN);

        byte[] red = new byte[3];
        red[0] = -1;
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getTopBorderXSSFColor().getRGB()).isEqualTo(red);
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getLeftBorderXSSFColor().getRGB()).isEqualTo(red);
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getBottomBorderXSSFColor().getRGB()).isEqualTo(red);
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getRightBorderXSSFColor().getRGB()).isEqualTo(red);
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getBorderTopEnum()).isEqualTo(BorderStyle.THICK);
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getBorderLeftEnum()).isEqualTo(BorderStyle.THICK);
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getBorderBottomEnum()).isEqualTo(BorderStyle.THICK);
        assertThat(xws.getRow(1).getCell(2).getCellStyle().getBorderRightEnum()).isEqualTo(BorderStyle.THICK);

        assertThat(xws.getRow(1).getCell(3).getCellStyle().getTopBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(3).getCellStyle().getLeftBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(3).getCellStyle().getBottomBorderXSSFColor().getRGB()).isEqualTo(red);
        assertThat(xws.getRow(1).getCell(3).getCellStyle().getRightBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(3).getCellStyle().getBorderTopEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(3).getCellStyle().getBorderLeftEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(3).getCellStyle().getBorderBottomEnum()).isEqualTo(BorderStyle.THICK);
        assertThat(xws.getRow(1).getCell(3).getCellStyle().getBorderRightEnum()).isEqualTo(BorderStyle.NONE);

        assertThat(xws.getRow(1).getCell(4).getCellStyle().getTopBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(4).getCellStyle().getLeftBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(4).getCellStyle().getBottomBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(4).getCellStyle().getRightBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(4).getCellStyle().getBorderTopEnum()).isEqualTo(BorderStyle.THIN);
        assertThat(xws.getRow(1).getCell(4).getCellStyle().getBorderLeftEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(4).getCellStyle().getBorderBottomEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(4).getCellStyle().getBorderRightEnum()).isEqualTo(BorderStyle.NONE);

        assertThat(xws.getRow(1).getCell(5).getCellStyle().getTopBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(5).getCellStyle().getLeftBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(5).getCellStyle().getBottomBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(5).getCellStyle().getRightBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(5).getCellStyle().getBorderTopEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(5).getCellStyle().getBorderLeftEnum()).isEqualTo(BorderStyle.THIN);
        assertThat(xws.getRow(1).getCell(5).getCellStyle().getBorderBottomEnum()).isEqualTo(BorderStyle.THICK);
        assertThat(xws.getRow(1).getCell(5).getCellStyle().getBorderRightEnum()).isEqualTo(BorderStyle.NONE);

        assertThat(xws.getRow(1).getCell(6).getCellStyle().getTopBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(6).getCellStyle().getLeftBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(6).getCellStyle().getBottomBorderXSSFColor()).isNull();
        assertThat(xws.getRow(1).getCell(6).getCellStyle().getRightBorderColor()).isEqualTo(IndexedColors.BLACK.index);
        assertThat(xws.getRow(1).getCell(6).getCellStyle().getBorderTopEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(6).getCellStyle().getBorderLeftEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(6).getCellStyle().getBorderBottomEnum()).isEqualTo(BorderStyle.NONE);
        assertThat(xws.getRow(1).getCell(6).getCellStyle().getBorderRightEnum()).isEqualTo(BorderStyle.THIN);

        for (int col = 1; col < 10; ++col) {
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getTopBorderXSSFColor()).isNull();
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getLeftBorderColor()).isEqualTo(IndexedColors.BLACK.index);
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getBottomBorderColor()).isEqualTo(IndexedColors.BLACK.index);
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getRightBorderXSSFColor()).isNull();
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getBorderTopEnum()).isEqualTo(BorderStyle.NONE);
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getBorderLeftEnum()).isEqualTo(BorderStyle.THIN);
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getBorderBottomEnum()).isEqualTo(BorderStyle.THIN);
            assertThat(xws.getRow(2).getCell(col).getCellStyle().getBorderRightEnum()).isEqualTo(BorderStyle.NONE);
        }
    }

    @Test
    public void listValidations() throws Exception {

        String errMsg = "Error Message";
        String errTitle = "Error Title";

        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");

            Worksheet listWs = wb.newWorksheet("Lists");
            listWs.value(0, 0, "val1");
            listWs.value(1, 0, "val2");

            Range listRange = listWs.range(0, 0, 1, 0);

            ListDataValidation listDataValidation = ws.range(0, 0, 100, 0).validateWithList(listRange);
            listDataValidation
                    .allowBlank(false)
                    .error(errMsg)
                    .errorTitle(errTitle)
                    .errorStyle(DataValidationErrorStyle.WARNING)
                    .showErrorMessage(true);
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertThat(xwb.getNumberOfSheets()).isEqualTo(2);
        XSSFSheet xws = xwb.getSheetAt(0);

        assertThat(xws.getDataValidations().size()).isEqualTo(1);

        XSSFDataValidation dataValidation = xws.getDataValidations().get(0);

        assertThat(dataValidation.getEmptyCellAllowed()).isFalse();
        assertThat(dataValidation.getErrorBoxText()).isEqualTo(errMsg);
        assertThat(dataValidation.getErrorBoxTitle()).isEqualTo(errTitle);
        assertThat(dataValidation.getErrorStyle()).isEqualTo(ErrorStyle.WARNING);
        assertThat(dataValidation.getShowErrorBox()).isTrue();
        assertThat(dataValidation.getSuppressDropDownArrow()).isTrue();
        assertThat(dataValidation.getRegions().getCellRangeAddresses().length).isEqualTo(1);

        CellRangeAddress cellRangeAddress = dataValidation.getRegions().getCellRangeAddress(0);
        assertThat(cellRangeAddress.getFirstColumn()).isEqualTo(0);
        assertThat(cellRangeAddress.getLastColumn()).isEqualTo(0);
        assertThat(cellRangeAddress.getFirstRow()).isEqualTo(0);
        assertThat(cellRangeAddress.getLastRow()).isEqualTo(100);

        DataValidationConstraint validationConstraint = dataValidation.getValidationConstraint();
        assertThat(validationConstraint.getFormula1().toLowerCase()).isEqualToIgnoringCase("Lists!A1:A2");
    }

    @Test
    public void canHideSheet() throws IOException {

        byte[] data = writeWorkbook(wb -> {
            wb.newWorksheet("Worksheet 1");
            Worksheet ws = wb.newWorksheet("Worksheet 2");
            ws.setVisibilityState(VisibilityState.HIDDEN);
            ws = wb.newWorksheet("Worksheet 3");
            ws.setVisibilityState(VisibilityState.VERY_HIDDEN);
            ws = wb.newWorksheet("Worksheet 4");
            ws.setVisibilityState(VisibilityState.VISIBLE);
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        assertThat(xwb.getSheetVisibility(0)).isEqualTo(SheetVisibility.VISIBLE);
        assertThat(xwb.getSheetVisibility(1)).isEqualTo(SheetVisibility.HIDDEN);
        assertThat(xwb.getSheetVisibility(2)).isEqualTo(SheetVisibility.VERY_HIDDEN);
        assertThat(xwb.getSheetVisibility(3)).isEqualTo(SheetVisibility.VISIBLE);
    }

    public void canHideColumns() throws Exception {

        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.hideColumn(1);
            ws.hideColumn(3);

            ws.value(0, 1, "val1");
            ws.value(0, 2, "val2");
            ws.value(0, 3, "val3");
            ws.value(0, 4, "val4");
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        XSSFSheet xws = xwb.getSheetAt(0);

        assertTrue("Column 1 should be hidden", xws.isColumnHidden(1));
        assertFalse("Column 2 should be visible", xws.isColumnHidden(2));
        assertTrue("Column 3 should be hidden", xws.isColumnHidden(3));
        assertFalse("Column 4 should be visible", xws.isColumnHidden(4));
    }

}
