package org.dhatim.fastexcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
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
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.dhatim.fastexcel.Correctness.writeWorkbook;
import static org.dhatim.fastexcel.SheetProtectionOption.*;
import static org.junit.Assert.*;

public class PoiCompatibility {

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
            ws.value(i, i++, Boolean.TRUE);
            ws.value(i, i++, Boolean.FALSE);
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
        Comparable<XSSFRow> row = xws.getRow(0);
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
        assertThat(xws.getRow(i).getCell(i++).getBooleanCellValue()).isTrue();
        assertThat(xws.getRow(i).getCell(i++).getBooleanCellValue()).isFalse();
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
        assertThat(IntStream.range(0, numWs).mapToObj(i -> xwb.getSheetName(i)))
                .containsExactly("Sheet 2", "Sheet 1", "Sheet 0");
    }

    @Test
    public void font() throws Exception {
        String sheetName = "Worksheet 1";
        String bold = "Bold";
        String italic = "Italic";
        String underlined = "Underlined";
        String bold_italic = "Bold_italic";
        String bold_underlined = "Bold_underlined";
        String italic_underlinded = "Italic_underlined";
        String all_three = "All_three";
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet(sheetName);
            ws.value(0, 0, bold);
            ws.style(0, 0).bold().set();
            ws.value(0, 1, italic);
            ws.style(0, 1).italic().set();
            ws.value(0, 2, underlined);
            ws.style(0, 2).underlined().set();
            ws.value(0, 3, bold_italic);
            ws.style(0, 3).bold().italic().set();
            ws.value(0, 4, bold_underlined);
            ws.style(0, 4).bold().underlined().set();
            ws.value(0, 5, italic_underlinded);
            ws.style(0, 5).italic().underlined().set();
            ws.value(0, 6, all_three);
            ws.style(0, 6).bold().italic().underlined().set();
            try {
                ws.finish();
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        XSSFSheet xws = xwb.getSheet(sheetName);
        assertTrue(xws.getRow(0).getCell(0).getCellStyle().getFont().getBold());
        assertTrue(xws.getRow(0).getCell(1).getCellStyle().getFont().getItalic());
        assertEquals(FontUnderline.valueOf(xws.getRow(0).getCell(2).getCellStyle().getFont().getUnderline()), FontUnderline.SINGLE);
        assertTrue(xws.getRow(0).getCell(3).getCellStyle().getFont().getBold());
        assertTrue(xws.getRow(0).getCell(3).getCellStyle().getFont().getItalic());
        assertTrue(xws.getRow(0).getCell(4).getCellStyle().getFont().getBold());
        assertEquals(FontUnderline.valueOf(xws.getRow(0).getCell(4).getCellStyle().getFont().getUnderline()), FontUnderline.SINGLE);
        assertTrue(xws.getRow(0).getCell(5).getCellStyle().getFont().getItalic());
        assertEquals(FontUnderline.valueOf(xws.getRow(0).getCell(5).getCellStyle().getFont().getUnderline()), FontUnderline.SINGLE);
        assertTrue(xws.getRow(0).getCell(6).getCellStyle().getFont().getBold());
        assertTrue(xws.getRow(0).getCell(6).getCellStyle().getFont().getItalic());
        assertEquals(FontUnderline.valueOf(xws.getRow(0).getCell(6).getCellStyle().getFont().getUnderline()), FontUnderline.SINGLE);
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
        assertThat(validationConstraint.getFormula1().toLowerCase()).isEqualToIgnoringCase("Lists!$A$1:$A$2");
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

    @Test
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

    @Test
    public void shouldHaveCorrectlyHashedPassword() throws IOException {

        final String password = "HorriblePassword";
        final String wrongPassword = "ThisIsNotThePassword";

        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.protect(password);
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        XSSFSheet xws = xwb.getSheetAt(0);

        // The valid password should be valid
        assertTrue(xws.validateSheetPassword(password));

        // The wrong password should be invalid
        assertFalse(xws.validateSheetPassword(wrongPassword));
    }

    @Test
    public void shouldSetAllDefaultSheetProtectionOptions() throws IOException {

        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.protect("HorriblePassword");
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        XSSFSheet xws = xwb.getSheetAt(0);

        // Only the defaults are locked.
        assertThat(xws.isAutoFilterLocked()).isTrue();
        assertThat(xws.isDeleteColumnsLocked()).isTrue();
        assertThat(xws.isDeleteRowsLocked()).isTrue();
        assertThat(xws.isFormatCellsLocked()).isTrue();
        assertThat(xws.isFormatColumnsLocked()).isTrue();
        assertThat(xws.isFormatRowsLocked()).isTrue();
        assertThat(xws.isInsertColumnsLocked()).isTrue();
        assertThat(xws.isInsertHyperlinksLocked()).isTrue();
        assertThat(xws.isInsertRowsLocked()).isTrue();
        assertThat(xws.isPivotTablesLocked()).isTrue();
        assertThat(xws.isSortLocked()).isTrue();
        assertThat(xws.isSheetLocked()).isTrue();

        // The rest is not locked
        assertThat(xws.isObjectsLocked()).isFalse();
        assertThat(xws.isScenariosLocked()).isFalse();
        assertThat(xws.isSelectLockedCellsLocked()).isFalse();
        assertThat(xws.isSelectUnlockedCellsLocked()).isFalse();
    }

    @Test
    public void shouldSetSpecificSheetProtectionOptions() throws IOException {

        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.protect("HorriblePassword", SHEET, DELETE_ROWS, OBJECTS);
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        XSSFSheet xws = xwb.getSheetAt(0);

        // Three are locked
        assertThat(xws.isSheetLocked()).isTrue();
        assertThat(xws.isDeleteRowsLocked()).isTrue();
        assertThat(xws.isObjectsLocked()).isTrue();

        // The rest is not locked
        assertThat(xws.isAutoFilterLocked()).isFalse();
        assertThat(xws.isDeleteColumnsLocked()).isFalse();
        assertThat(xws.isFormatCellsLocked()).isFalse();
        assertThat(xws.isFormatColumnsLocked()).isFalse();
        assertThat(xws.isFormatRowsLocked()).isFalse();
        assertThat(xws.isInsertColumnsLocked()).isFalse();
        assertThat(xws.isInsertHyperlinksLocked()).isFalse();
        assertThat(xws.isInsertRowsLocked()).isFalse();
        assertThat(xws.isPivotTablesLocked()).isFalse();
        assertThat(xws.isSortLocked()).isFalse();
        assertThat(xws.isScenariosLocked()).isFalse();
        assertThat(xws.isSelectLockedCellsLocked()).isFalse();
        assertThat(xws.isSelectUnlockedCellsLocked()).isFalse();
    }

    @Test
    public void canProtectCells() throws Exception {

        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");

            // Protect the sheet
            ws.protect("HorriblePassword");

            ws.value(0, 1, "val1");
            ws.style(0, 1)
                    .protectionOption(ProtectionOption.LOCKED, true).set();

            ws.value(0, 2, "val2");
            ws.style(0, 2)
                    .protectionOption(ProtectionOption.HIDDEN, true).set();

            ws.value(0, 3, "val3");
            ws.style(0, 3)
                    .protectionOption(ProtectionOption.LOCKED, false).set();

            ws.value(0, 4, "val4");
            ws.style(0, 4)
                    .protectionOption(ProtectionOption.LOCKED, false)
                    .protectionOption(ProtectionOption.HIDDEN, false).set();
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        XSSFSheet xws = xwb.getSheetAt(0);

        assertFalse("Cell (0, 1) should NOT be hidden", xws.getRow(0).getCell(1).getCellStyle().getHidden());
        assertTrue("Cell (0, 1) should be locked", xws.getRow(0).getCell(1).getCellStyle().getLocked());

        assertTrue("Cell (0, 2) should be locked (by default)", xws.getRow(0).getCell(2).getCellStyle().getLocked());
        assertTrue("Cell (0, 2) should be hidden", xws.getRow(0).getCell(2).getCellStyle().getHidden());

        assertFalse("Cell (0, 3) should NOT be hidden", xws.getRow(0).getCell(3).getCellStyle().getHidden());
        assertFalse("Cell (0, 3) should NOT be locked", xws.getRow(0).getCell(3).getCellStyle().getLocked());

        assertFalse("Cell (0, 4) should NOT be locked", xws.getRow(0).getCell(4).getCellStyle().getLocked());
        assertFalse("Cell (0, 4) should NOT be hidden", xws.getRow(0).getCell(4).getCellStyle().getHidden());
    }

    @Test
    public void comments() throws IOException {
        final String commentText01 = "this is a comment <&>";
        final String commentText32 = "another comment";
        byte[] data = writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 1, "cell value");
            ws.comment(0, 1, commentText01);
            ws.comment(3, 2, commentText32);
        });

        // Check generated workbook with Apache POI
        XSSFWorkbook xwb = new XSSFWorkbook(new ByteArrayInputStream(data));
        XSSFSheet xws = xwb.getSheetAt(0);
        assertEquals(
                "B1: this is a comment <&>\n" +
                        "C4: another comment",
                xws.getCellComments().entrySet().stream()
                        .map(e -> e.getKey() + ": " + e.getValue().getString())
                        .collect(Collectors.joining("\n"))
        );
        XSSFComment cellComment = xws.getRow(0).getCell(1).getCellComment();
        assertNotNull("Comment should NOT be null", cellComment);
        assertEquals(
                commentText01,
                String.valueOf(cellComment.getString())
        );
    }
}
