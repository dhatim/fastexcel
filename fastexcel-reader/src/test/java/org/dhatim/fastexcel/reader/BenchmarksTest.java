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

import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.SAXHelper;
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

public class BenchmarksTest {

    private static class SheetContentHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

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
            //reading
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
            //Do nothing
        }
    }

    private static final String FILE = "/xlsx/calendar_stress_test.xlsx";

    @Benchmark
    public void apachePoi() throws IOException, InvalidFormatException {
        try (Workbook wb = WorkbookFactory.create(openResource(FILE))) {
            org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(0);
            sheet.forEach(r -> {
                r.getCell(0);
            });
        }
    }

    @Benchmark
    public void fastExcelReader() throws IOException {
        try (InputStream is = openResource(FILE); ReadableWorkbook wb = new ReadableWorkbook(is)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(r -> {
                    Cell cell = r.getCell(0);
                });
            }
        }
    }

    @Benchmark
    public void streamingApachePoi() throws IOException, OpenXML4JException, SAXException {
        try (OPCPackage pkg = OPCPackage.open(openResource(FILE))) {
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            XSSFReader reader = new XSSFReader(pkg);
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) reader.getSheetsData();
            int sheetIndex = 0;
            while (iterator.hasNext()) {
                try (InputStream sheetStream = iterator.next()) {
                    if (sheetIndex == 0) {
                        processSheet(styles, strings, new SheetContentHandler(), sheetStream);
                    }
                }
                sheetIndex++;
            }
        }
    }

    private void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings,
                              XSSFSheetXMLHandler.SheetContentsHandler sheetHandler, InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch(ParserConfigurationException e) {
            throw new IllegalStateException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    @Test
    public void benchmarks() throws RunnerException {
        String foo = BenchmarksTest.class.getName() + "\\..*";
        Options options = new OptionsBuilder().include(foo)
                .mode(Mode.SingleShotTime)
                .warmupIterations(0)
                .warmupBatchSize(1)
                .measurementIterations(1)
                .threads(1)
                .forks(0)
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
