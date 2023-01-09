package org.dhatim.fastexcel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.security.GeneralSecurityException;

public class ProtectWriteTest {
    @Test
    void fastexcelWriteProtectTest() throws IOException, GeneralSecurityException, InvalidFormatException {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream(); POIFSFileSystem fs = new POIFSFileSystem()) {
            Workbook wb = new Workbook(bos, "Test", "1.0");
            wb.setGlobalDefaultFont("Arial", 15.5);
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, "Hello fastexcel");
            wb.finish();
            byte[] bytes = bos.toByteArray();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            // EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);
            Encryptor enc = info.getEncryptor();
            enc.confirmPassword("1234");
            // Read in an existing OOXML file and write to encrypted output stream
            // don't forget to close the output stream otherwise the padding bytes aren't added
            try (OPCPackage opc = OPCPackage.open(new ByteArrayInputStream(bytes)); OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
            }
            // Write out the encrypted version
            try (FileOutputStream fos = new FileOutputStream("D://protectedTest.xlsx")) {
                fs.writeFilesystem(fos);
            }
        }

    }
}
