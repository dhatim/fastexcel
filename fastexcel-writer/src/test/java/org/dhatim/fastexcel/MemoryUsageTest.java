package org.dhatim.fastexcel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.junit.Test;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.logging.Logger;

import static org.junit.Assert.assertEquals;

public class MemoryUsageTest {
    private static final Logger LOG = Logger.getLogger(MemoryUsageTest.class.getName());
    private static final int ROWS = 500_000;
    private static final int COLS = 100;
    private static final int SHEETS = 2;
    private static final File FILE = new File("target/memtest" + ROWS + "x" + COLS + ".xlsx");

    private class SheetContentHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private NumberFormat FORMAT = new DecimalFormat("0");
        public int processedRows = 0;
        public int processedCells = 0;
        private int sheetIndex;

        public SheetContentHandler(int sheetIndex) {
            this.sheetIndex = sheetIndex;
        }

        @Override
        public void startRow(int rowNum) {
            printProgress("validating", sheetIndex * ROWS + rowNum, SHEETS * ROWS);
        }

        @Override
        public void endRow(int rowNum) {
            processedRows++;
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            CellReference ref = new CellReference(cellReference);
            assertEquals(FORMAT.format(valueFor(ref.getRow(), ref.getCol())), formattedValue);
            processedCells++;
        }
    }

    @Test
    public void run() throws Exception {
        try (OutputStream out = new FileOutputStream(FILE)) {
            generate(out);
        }
        validate();
    }

    private void validate() throws Exception {
        try (OPCPackage pkg = OPCPackage.open(FILE)) {
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            XSSFReader reader = new XSSFReader(pkg);
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) reader.getSheetsData();
            int sheetIndex = 0;
            while (iterator.hasNext()) {
                try (InputStream sheetStream = iterator.next()) {
                    SheetContentHandler handler = new SheetContentHandler(sheetIndex);
                    processSheet(styles, strings, handler, sheetStream);
                    assertEquals(ROWS, handler.processedRows);
                    assertEquals(ROWS * COLS, handler.processedCells);
                }
                sheetIndex++;
            }
            assertEquals(SHEETS, sheetIndex);
        }
    }

    private void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings,
                              XSSFSheetXMLHandler.SheetContentsHandler sheetHandler, InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        saxFactory.setNamespaceAware(true);
        try {
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader sheetParser = saxParser.getXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    public void generate(OutputStream out) throws IOException {
        Workbook wb = new Workbook(out, "test", "1.0");
        for (int s = 0; s < SHEETS; s++) {
            Worksheet sheet = wb.newWorksheet("sheet " + s);
            for (int r = 0; r < ROWS; r++) {
                printProgress("writing", s * ROWS + r, ROWS * SHEETS);
                for (int c = 0; c < COLS; c++) {
                    sheet.value(r, c, valueFor(r, c));
                }
                if (r % 100 == 0) {
                    sheet.flush();
                }
            }
            sheet.finish();
        }
        wb.finish();
    }

    private static double valueFor(int r, int c) {
        return (double) r * COLS + c;
    }

    private static void printProgress(String prefix, int r, int total) {
        if (r % (total / 100) == 0) {
            LOG.info(prefix + ": " + (100 * r / total) + "%");
        }
    }
}
