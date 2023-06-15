package org.dhatim.fastexcel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.assertEquals;


class MarginalInformationTest {

    @Test
    public void testMarginalInformationMinimal() {
        MarginalInformation minimal = new MarginalInformation("minimal-example", Position.CENTER);
        assertEquals("&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;K000000minimal-example", minimal.getContent());
    }

    @Test
    public void testMarginalInformationFont() {
        MarginalInformation font = new MarginalInformation("font-example", Position.CENTER).withFont("Arial");
        assertEquals("&amp;C&amp;&quot;Arial,Regular&quot;&amp;12&amp;K000000font-example", font.getContent());
    }

    @Test
    public void testMarginalInformationFontSize() {
        MarginalInformation fontSize = new MarginalInformation("font-example", Position.CENTER).withFontSize(9);
        assertEquals("&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;9&amp;K000000font-example", fontSize.getContent());
    }

    @Test
    public void testMarginalInformationFull() {
        MarginalInformation full = new MarginalInformation("full-example", Position.CENTER).withFont("Arial").withFontSize(9);
        assertEquals("&amp;C&amp;&quot;Arial,Regular&quot;&amp;9&amp;K000000full-example", full.getContent());
    }

    @Test
    public void testMarginalInformationEscaping() {
        MarginalInformation full = new MarginalInformation("es'ca&pi<ng>-example", Position.CENTER);
        assertEquals("&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;K000000es&apos;ca&amp;pi&lt;ng&gt;-example", full.getContent());
    }

    @ParameterizedTest
    @CsvSource(value = {
        "page 1 of ? : &amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;K000000Page &amp;P of &amp;N",
        "page 1, sheetname : &amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;K000000Page &amp;P, &amp;A",
        "page 1 : &amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;K000000Page &amp;P",
        "sheetname : &amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;K000000&amp;A",

    }, delimiter = ':')
    public void testMarginalInformationTemplates(String input, String expected) {
        MarginalInformation info = new MarginalInformation(input, Position.CENTER);
        assertEquals(expected, info.getContent());
    }


    // Regression test for Github issue #302 (https://github.com/dhatim/fastexcel/issues/302)
    @Test
    void testHeaderEscapingIntegration() throws Exception {
        File tempFile = File.createTempFile("fastexcel-", "");
        tempFile.deleteOnExit();
        FileOutputStream os = new FileOutputStream(tempFile);
        try (Workbook wb = new Workbook(os, "MyApp", "0.0")) {
            Worksheet ws = wb.newWorksheet("Sheet 1");
            ws.header("<Left> & 'left' & \"LEFT\"", Position.LEFT, "Calibri", 12);
            ws.header("<Center> & 'center' & \"CENTER\"", Position.CENTER, "Calibri", 12);
            ws.header("<Right> & 'right' & \"RIGHT\"", Position.RIGHT, "Calibri", 12);
            ws.close();
        }
        // use POI for integration test as a 3rd party reference
        // check if footer is readable, if the text is not correctly escaped, it will fail
        XSSFWorkbook referenceWorkbook = new XSSFWorkbook(Files.newInputStream(tempFile.toPath()));
        XSSFSheet sheet = referenceWorkbook.getSheet("Sheet 1");
        String centerHeader = sheet.getHeader().getCenter();
        assertEquals("&\"Calibri,Regular\"&12&K000000<Center> & 'center' & \"CENTER\"", centerHeader);
    }

    // Regression test for Github issue #302 (https://github.com/dhatim/fastexcel/issues/302)
    @Test
    void testFooterEscapingIntegration() throws Exception {
        File tempFile = File.createTempFile("fastexcel-", "");
        tempFile.deleteOnExit();
        FileOutputStream os = new FileOutputStream(tempFile);
        try (Workbook wb = new Workbook(os, "MyApp", "0.0")) {
            Worksheet ws = wb.newWorksheet("Sheet 1");
            ws.footer("<Left> & 'left' & \"LEFT\"", Position.LEFT, "Calibri", 12);
            ws.footer("<Center> & 'center' & \"CENTER\"", Position.CENTER, "Calibri", 12);
            ws.footer("<Right> & 'right' & \"RIGHT\"", Position.RIGHT, "Calibri", 12);
            ws.close();
        }
        // use POI for integration test as a 3rd party reference
        // check if footer is readable, if the text is not correctly escaped, it will fail
        XSSFWorkbook referenceWorkbook = new XSSFWorkbook(Files.newInputStream(tempFile.toPath()));
        XSSFSheet sheet = referenceWorkbook.getSheet("Sheet 1");
        String centerFooter = sheet.getFooter().getCenter();
        assertEquals("&\"Calibri,Regular\"&12&K000000<Center> & 'center' & \"CENTER\"", centerFooter);
    }
}
