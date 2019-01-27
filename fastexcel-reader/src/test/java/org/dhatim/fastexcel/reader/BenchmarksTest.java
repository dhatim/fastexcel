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

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.junit.jupiter.api.Test;
import org.openjdk.jmh.annotations.Benchmark;
import org.openjdk.jmh.annotations.Mode;
import org.openjdk.jmh.results.format.ResultFormatType;
import org.openjdk.jmh.runner.Runner;
import org.openjdk.jmh.runner.RunnerException;
import org.openjdk.jmh.runner.options.Options;
import org.openjdk.jmh.runner.options.OptionsBuilder;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class BenchmarksTest {

    private static final long RESULT = 2147385345;

    private static class SheetContentHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        private long result = 0;

        @Override
        public void startRow(int rowNum) {
            //Do nothing
        }

        @Override
        public void endRow(int rowNum) {
            //Do nothing
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            CellReference ref = new CellReference(cellReference);
            if (ref.getRow() > 0 && ref.getCol() == 0) {
                result += Double.parseDouble(formattedValue);
            }
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
            //Do nothing
        }
    }

    private static final String FILE = "/xlsx/calendar_stress_test.xlsx";

    @Benchmark
    public long apachePoi() throws IOException {
        try (Workbook wb = WorkbookFactory.create(openResource(FILE))) {
            org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(0);
            long sum = StreamSupport.stream(
                    Spliterators.spliteratorUnknownSize(sheet.rowIterator(), Spliterator.ORDERED),
                    false
            ).skip(1).mapToLong(r -> (long) r.getCell(0).getNumericCellValue()).sum();
            assertEquals(RESULT, sum);
            return sum;
        }
    }

    @Benchmark
    public long fastExcelReader() throws IOException {
        try (InputStream is = openResource(FILE); ReadableWorkbook wb = new ReadableWorkbook(is)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                long sum = rows.skip(1).mapToLong(r -> r.getCell(0).asNumber().longValue()).sum();
                assertEquals(RESULT, sum);
                return sum;
            }
        }
    }

    @Benchmark
    public int streamingApachePoi() throws IOException, OpenXML4JException, SAXException {
        try (OPCPackage pkg = OPCPackage.open(openResource(FILE))) {
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            XSSFReader reader = new XSSFReader(pkg);
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) reader.getSheetsData();
            int sheetIndex = 0;
            while (iterator.hasNext()) {
                try (InputStream sheetStream = iterator.next()) {
                    if (sheetIndex == 0) {
                        SheetContentHandler sheetHandler = new SheetContentHandler();
                        processSheet(styles, strings, sheetHandler, sheetStream);
                        assertEquals(RESULT, sheetHandler.result);
                    }
                }
                sheetIndex++;
            }
            return sheetIndex;
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


    @Benchmark
    public long monitorjbl() throws IOException {
        long sum = 0;
        try (InputStream is = openResource(FILE);
             org.apache.poi.ss.usermodel.Workbook workbook = StreamingReader.builder().open(is)) {
            for (org.apache.poi.ss.usermodel.Sheet sheet : workbook) {
                for (org.apache.poi.ss.usermodel.Row r : sheet) {
                    if (r.getRowNum() == 0) {
                        continue;
                    }
                    sum += r.getCell(0).getNumericCellValue();
                }
            }
        }
        assertEquals(RESULT, sum);
        return sum;
    }

    @Test
    public void benchmarks() throws RunnerException {
        String foo = BenchmarksTest.class.getName() + "\\..*";
        Options options = new OptionsBuilder().include(foo)
                .mode(Mode.SingleShotTime)
                .warmupIterations(1)
                .warmupBatchSize(1)
                .measurementIterations(1)
                .threads(1)
                .forks(1)
                .timeUnit(TimeUnit.MILLISECONDS)
                .shouldFailOnError(true)
                .resultFormat(ResultFormatType.CSV)
                .result("jmh.csv")
                .build();
        new Runner(options).run();
    }

    private static InputStream openResource(String name) {
        InputStream result = BenchmarksTest.class.getResourceAsStream(name);
        if (result == null) {
            throw new IllegalStateException("Cannot read resource " + name);
        }
        return result;
    }

}
