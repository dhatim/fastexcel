package org.dhatim.fastexcel.reader;

import java.io.Closeable;
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
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;

public class ReadableWorkbook implements Closeable {

    private final OPCPackage pkg;
    private final XSSFReader reader;
    private final SharedStringsTable sst;
    private final XMLInputFactory factory;

    private boolean date1904;
    private final List<Sheet> sheets = new ArrayList<>();

    public ReadableWorkbook(InputStream inputStream) throws IOException {
        try {
            pkg = OPCPackage.open(inputStream);
            reader = new XSSFReader(pkg);
            sst = reader.getSharedStringsTable();
        } catch (OpenXML4JException e) {
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
                date1904 = date1904Value == null ? false : Boolean.parseBoolean(date1904Value);
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

    private static Runnable asUncheckedRunnable(Closeable c) {
        return () -> {
            try {
                c.close();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        };
    }

}
