package org.dhatim.fastexcel;

import com.github.rzymek.opczip.reader.skipping.ZipStreamReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class MacroCopyTest {

    @Test
    void testCopyMacro() throws IOException {
        byte[] document;
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            try (Workbook workbook = new Workbook(out, "FastExcel", "1.0")) {
                Worksheet sheet = workbook.newWorksheet("Hello");
                sheet.value(1, 1, "Hello World");
                workbook.copyMacrosFromInputStream(MacroCopyTest.class.getResourceAsStream("HelloWorldMacro.xlsm"));
            }
            out.flush();
            document = out.toByteArray();
        }

        boolean hasVbaProjectEmbedded = false;
        try (ZipStreamReader zipInputStream = new ZipStreamReader(new ByteArrayInputStream(document))) {
            ZipEntry entry;
            while ((entry = zipInputStream.nextEntry()) != null) {
                if (entry.getName().equals("xl/vbaProject.bin")) {
                    hasVbaProjectEmbedded = true;
                    break;
                }
                zipInputStream.skipStream();
                zipInputStream.skipStream();
            }
        }
        assertTrue(hasVbaProjectEmbedded);
    }
}
