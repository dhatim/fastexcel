package org.dhatim.fastexcel.reader;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;

public class ProtectedReadTest {
  @Test
    void fastexcelReadProtectTest() throws IOException {
        try (POIFSFileSystem fileSystem = new POIFSFileSystem(new File("D://protectedTest.xlsx"))){
            EncryptionInfo info = new EncryptionInfo(fileSystem);
            Decryptor d = Decryptor.getInstance(info);
            if (!d.verifyPassword("1234")) {
                throw new RuntimeException("Unable to process: document is encrypted");
            }
            // parse dataStream
            try (InputStream dataStream = d.getDataStream(fileSystem); ReadableWorkbook fworkbook = new ReadableWorkbook(dataStream)) {
                Sheet sheet = fworkbook.getSheet(0).orElse(null);
                assert sheet != null;
                sheet.openStream().forEach(System.out::println);
            }
        } catch (GeneralSecurityException ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        }
    }


}
