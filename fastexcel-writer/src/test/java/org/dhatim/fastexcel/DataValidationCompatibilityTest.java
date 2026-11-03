package org.dhatim.fastexcel;

import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;

class DataValidationCompatibilityTest {

    @ParameterizedTest
    @CsvSource({"list, true", "list, false", "listFormula, true", "listFormula, false",
            "custom, true", "custom, false"})
    void errorAlertsWithXmlCharactersRoundTrip(String type, boolean escapeTitle) throws Exception {
        String specialText = "Choose \"A & B\" <or> 'C' — café";
        String title = escapeTitle ? specialText : "Invalid value";
        String message = escapeTitle ? "Choose a valid value" : specialText;
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (Workbook workbook = new Workbook(output, "Test", "1.0")) {
            Worksheet sheet = workbook.newWorksheet("Sheet1");
            sheet.value(0, 1, "A");
            sheet.value(1, 1, "B");
            Range target = sheet.range(0, 0, 1, 0);
            switch (type) {
                case "list":
                    target.validateWithList(sheet.range(0, 1, 1, 1))
                            .errorTitle(title).error(message).showErrorMessage(true);
                    break;
                case "listFormula":
                    target.validateWithListByFormula("$B$1:$B$2")
                            .errorTitle(title).error(message).showErrorMessage(true);
                    break;
                case "custom":
                    target.validateWithFormula("ISNUMBER(A1)")
                            .errorTitle(title).error(message).showErrorMessage(true);
                    break;
                default:
                    throw new IllegalArgumentException(type);
            }
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            assertThat(workbook.getSheetAt(0).getDataValidations()).hasSize(1);
            XSSFDataValidation validation = workbook.getSheetAt(0).getDataValidations().get(0);
            assertThat(validation.getErrorBoxTitle()).isEqualTo(title);
            assertThat(validation.getErrorBoxText()).isEqualTo(message);
            assertThat(validation.getShowErrorBox()).isTrue();
        }
    }
}
