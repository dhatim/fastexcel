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

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.common.POIFSConstants;
import org.apache.poi.poifs.storage.HeaderBlockConstants;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;

public class ReadableWorkbook implements Closeable {

    private final OPCPackage pkg;
    private final XSSFReader reader;
    private final SharedStringsTable sst;
    private final XMLInputFactory factory;

    private boolean date1904;
    private final List<Sheet> sheets = new ArrayList<>();

    public ReadableWorkbook(File inputFile) throws IOException {
        this(open(inputFile));
    }

    /**
     * Note: when working with huge sheets (e.g. 500_000 rows) use {@link #ReadableWorkbook(File)}
     */
    public ReadableWorkbook(InputStream inputStream) throws IOException {
        this(open(inputStream));
    }

    private ReadableWorkbook(OPCPackage pkg) throws IOException {
        try {
            this.pkg = pkg;
            reader = new XSSFReader(pkg);
            sst = reader.getSharedStringsTable();
        } catch (NotOfficeXmlFileException | OpenXML4JException e) {
            throw new ExcelReaderException(e);
        }
        factory = XMLInputFactory.newInstance();
        // To prevent XML External Entity (XXE) attacks
        factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, Boolean.FALSE);

        try (SimpleXmlReader workbookReader = new SimpleXmlReader(factory, reader.getWorkbookData())) {
            readWorkbook(workbookReader);
        } catch (InvalidFormatException | XMLStreamException e) {
            throw new ExcelReaderException(e);
        }
    }

    @Override
    public void close() throws IOException {
        pkg.close();
    }

    public boolean isDate1904() {
        return date1904;
    }

    public Stream<Sheet> getSheets() {
        return sheets.stream();
    }

    public Optional<Sheet> getSheet(int index) {
        return index < 0 || index >= sheets.size() ? Optional.empty() : Optional.of(sheets.get(index));
    }

    public Sheet getFirstSheet() {
        return sheets.get(0);
    }

    public Optional<Sheet> findSheet(String name) {
        return sheets.stream().filter(sheet -> name.equals(sheet.getName())).findFirst();
    }

    private void readWorkbook(SimpleXmlReader r) throws XMLStreamException {
        while (r.goTo(() -> r.isStartElement("sheets") || r.isStartElement("workbookPr") || r.isEndElement("workbook"))) {
            if ("sheets".equals(r.getLocalName())) {
                r.forEach("sheet", "sheets", this::createSheet);
            } else if ("workbookPr".equals(r.getLocalName())) {
                String date1904Value = r.getAttribute("date1904");
                date1904 = Boolean.parseBoolean(date1904Value);
            } else {
                break;
            }
        }
    }

    private void createSheet(SimpleXmlReader r) {
        String name = r.getAttribute("name");
        String id = r.getAttribute("id");
        int index = sheets.size();
        sheets.add(new Sheet(this, index, id, name));
    }

    Stream<Row> openStream(Sheet sheet) throws IOException {
        try {
            InputStream inputStream = reader.getSheet(sheet.getId());
            Stream<Row> stream = StreamSupport.stream(new RowSpliterator(this, inputStream), false);
            return stream.onClose(asUncheckedRunnable(inputStream));
        } catch (XMLStreamException | InvalidFormatException e) {
            throw new IOException(e);
        }
    }

    XMLInputFactory getXmlFactory() {
        return factory;
    }

    SharedStringsTable getSharedStringsTable() {
        return sst;
    }

    public static boolean isOOXMLZipHeader(byte[] bytes) {
        requireLength(bytes, POIFSConstants.OOXML_FILE_HEADER.length);
        return arrayEquals(bytes, 0, POIFSConstants.OOXML_FILE_HEADER, 0, POIFSConstants.OOXML_FILE_HEADER.length);
    }

    public static boolean isOLE2Header(byte[] bytes) {
        requireLength(bytes, 8);
        byte[] ole2Header = new byte[8];
        LittleEndian.putLong(ole2Header, 0, HeaderBlockConstants._signature);
        return arrayEquals(bytes, 0, ole2Header, 0, ole2Header.length);
    }

    private static void requireLength(byte[] bytes, int requiredLength) {
        if (bytes.length < requiredLength) {
            throw new IllegalArgumentException("Insufficient header bytes");
        }
    }

    private static boolean arrayEquals(byte[] a, int offsetA, byte[] b, int offsetB, int length) {
        if (a == b) {
            return true;
        }
        if (a == null || b == null) {
            return false;
        }
        if ((offsetA + length > a.length) || (offsetB + length > b.length)) {
            return false;
        }
        for (int i = 0; i < length; i++) {
            if (a[offsetA + i] != b[offsetB + i]) {
                return false;
            }
        }
        return true;
    }

    private static Runnable asUncheckedRunnable(Closeable c) {
        return () -> {
            try {
                c.close();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        };
    }

    private static OPCPackage open(File file){
        try {
            return OPCPackage.open(file);
        } catch (InvalidFormatException e) {
            throw new ExcelReaderException(e);
        }
    }

    private static OPCPackage open(InputStream in) throws IOException {
        try {
            return OPCPackage.open(in);
        } catch (InvalidFormatException e) {
            throw new ExcelReaderException(e);
        }
    }

}
