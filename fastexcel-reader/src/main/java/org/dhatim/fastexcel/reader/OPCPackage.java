package org.dhatim.fastexcel.reader;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.commons.compress.utils.SeekableInMemoryByteChannel;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

class OPCPackage implements AutoCloseable {
    private final ZipFile zip;

    private OPCPackage(File zipFile) throws IOException {
        zip = new ZipFile(zipFile);
    }

    OPCPackage(SeekableInMemoryByteChannel channel) throws IOException {
        zip = new ZipFile(channel);
    }

    static OPCPackage open(File inputFile) throws IOException {
        return new OPCPackage(inputFile);
    }

    static OPCPackage open(InputStream inputStream) throws IOException {
        byte[] compressedBytes = IOUtils.toByteArray(inputStream);
        return new OPCPackage(new SeekableInMemoryByteChannel(compressedBytes));
    }

    InputStream getSharedStrings() throws IOException {
        return getEntryContent("xl/sharedStrings.xml");
    }


    private InputStream getEntryContent(String name) throws IOException {
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

    public InputStream getWorkbook() throws IOException {
        return getEntryContent("xl/workbook.xml");
    }

    public InputStream getSheetContent(Sheet sheet) throws IOException {
        String name = "xl/worksheets/sheet" + (sheet.getIndex() + 1) + ".xml";
        InputStream inputStream = getEntryContent(name);
        if (inputStream == null) {
            throw new IOException(name + " not found");
        }
        return inputStream;
    }
}
