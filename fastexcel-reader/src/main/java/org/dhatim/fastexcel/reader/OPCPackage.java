package org.dhatim.fastexcel.reader;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.commons.compress.utils.SeekableInMemoryByteChannel;

import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Pattern;

import static java.lang.String.format;
import static org.dhatim.fastexcel.reader.DefaultXMLInputFactory.factory;

class OPCPackage implements AutoCloseable {
    private static final Pattern filenameRegex = Pattern.compile("^(.*/)([^/]+)$");
    private final ZipFile zip;
    private final Map<String, String> workbookPartsById;
    private final PartEntryNames parts;

    private OPCPackage(File zipFile) throws IOException {
        this(new ZipFile(zipFile));
    }

    private OPCPackage(SeekableInMemoryByteChannel channel) throws IOException {
        this(new ZipFile(channel));
    }

    private OPCPackage(ZipFile zip) throws IOException {
        try {
            this.zip = zip;
            this.parts = extractPartEntriesFromContentTypes();
            this.workbookPartsById = readWorkbookPartsIds(relsNameFor(parts.workbook));
        } catch (XMLStreamException e) {
            throw new IOException(e);
        }
    }

    private static String relsNameFor(String entryName) {
        return filenameRegex.matcher(entryName).replaceFirst("$1_rels/$2.rels");
    }

    private Map<String, String> readWorkbookPartsIds(String workbookRelsEntryName) throws IOException, XMLStreamException {
        Map<String, String> partsIdById = new HashMap<>();
        SimpleXmlReader rels = new SimpleXmlReader(factory, getRequiredEntryContent(workbookRelsEntryName));
        while (rels.goTo("Relationship")) {
            String id = rels.getAttribute("Id");
            String target = rels.getAttribute("Target");
            partsIdById.put(id, target);
        }
        return partsIdById;
    }

    private PartEntryNames extractPartEntriesFromContentTypes() throws XMLStreamException, IOException {
        PartEntryNames entries = new PartEntryNames();
        final String contentTypesXml = "[Content_Types].xml";
        try (SimpleXmlReader reader = new SimpleXmlReader(factory, getRequiredEntryContent(contentTypesXml))) {
            while (reader.goTo(() -> reader.isStartElement("Override"))) {
                if (PartEntryNames.WORKBOOK_MAIN_CONTENT_TYPE.equals(reader.getAttributeRequired("ContentType"))) {
                    entries.workbook = reader.getAttributeRequired("PartName");
                }
                if (PartEntryNames.SHARED_STRINGS_CONTENT_TYPE.equals(reader.getAttributeRequired("ContentType"))) {
                    entries.sharedStrings = reader.getAttributeRequired("PartName");
                }
                if (entries.isFullyFilled()) {
                    break;
                }
            }
        }
        return entries;
    }

    private InputStream getRequiredEntryContent(String name) throws IOException {
        return Optional.ofNullable(getEntryContent(name))
            .orElseThrow(() -> new ExcelReaderException(name + " not found"));
    }

    static OPCPackage open(File inputFile) throws IOException {
        return new OPCPackage(inputFile);
    }

    static OPCPackage open(InputStream inputStream) throws IOException {
        byte[] compressedBytes = IOUtils.toByteArray(inputStream);
        return new OPCPackage(new SeekableInMemoryByteChannel(compressedBytes));
    }

    InputStream getSharedStrings() throws IOException {
        return getEntryContent(parts.sharedStrings);
    }

    private InputStream getEntryContent(String name) throws IOException {
        if (name.startsWith("/")) {
            name = name.substring(1);
        }
        ZipArchiveEntry entry = zip.getEntry(name);
        if (entry == null) {
            return null;
        }
        return zip.getInputStream(entry);
    }

    @Override
    public void close() throws IOException {
        zip.close();
    }

    public InputStream getWorkbookContent() throws IOException {
        return getRequiredEntryContent(parts.workbook);
    }

    public InputStream getSheetContent(Sheet sheet) throws IOException {
        String name = this.workbookPartsById.get(sheet.getId());
        if (name == null) {
            String msg = format("Sheet#%s '%s' is missing an entry in workbook rels (for id: '%s')",
                sheet.getIndex(), sheet.getName(), sheet.getId());
            throw new ExcelReaderException(msg);
        }
        return getRequiredEntryContent("xl/" + name);
    }

    private static class PartEntryNames {
        public static final String WORKBOOK_MAIN_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        public static final String SHARED_STRINGS_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        String workbook;
        String sharedStrings;

        boolean isFullyFilled() {
            return workbook != null && sharedStrings != null;
        }
    }
}
