package org.dhatim.fastexcel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Sheet;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.security.GeneralSecurityException;

public class EncryptionTest {

    private static final File testFile = new File("target/encryptTest.xlsx");

    private static final String secretKey = "foobaa";

    private static final String testContent = "Hello fastexcel";


    void fastexcelWriteProtectTest() throws IOException, GeneralSecurityException, InvalidFormatException {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream(); POIFSFileSystem fs = new POIFSFileSystem()) {
            Workbook wb = new Workbook(bos, "Test", "1.0");
            wb.setGlobalDefaultFont("Arial", 15.5);
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, testContent);
            wb.finish();
            byte[] bytes = bos.toByteArray();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            // EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);
            Encryptor enc = info.getEncryptor();
            enc.confirmPassword(secretKey);
            // Read in an existing OOXML file and write to encrypted output stream
            // don't forget to close the output stream otherwise the padding bytes aren't added
            try (OPCPackage opc = OPCPackage.open(new ByteArrayInputStream(bytes)); OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
            }
            // Write out the encrypted version
            try (FileOutputStream fos = new FileOutputStream(testFile)) {
                fs.writeFilesystem(fos);
            }
        }

    }

    void fastexcelReadProtectTest() throws IOException, InvalidFormatException {
        try (POIFSFileSystem fileSystem = new POIFSFileSystem(testFile)) {
            EncryptionInfo info = new EncryptionInfo(fileSystem);
            Decryptor d = Decryptor.getInstance(info);
            if (!d.verifyPassword(secretKey)) {
                throw new RuntimeException("Unable to process: document is encrypted");
            }
            // parse dataStream
            try (InputStream dataStream = d.getDataStream(fileSystem); ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                /* TODO:The dataStream obtained here is broken and cannot be read normally by programs other than POI, which is probably a bug of POI.
                     See https://bz.apache.org/bugzilla/show_bug.cgi?id=66436
                     This problem can be avoided by saving after being wrapped by OPCPackage */
                OPCPackage open = OPCPackage.open(dataStream);
                open.save(bos);
                byte[] bytes = bos.toByteArray();
                ReadableWorkbook fworkbook = new ReadableWorkbook(new ByteArrayInputStream(bytes));
                Sheet sheet = fworkbook.getSheet(0).orElse(null);
                assert sheet != null;
                sheet.openStream().forEach(r -> {
                    r.forEach(c -> System.out.println(c.getText()));
                });
            }
        } catch (GeneralSecurityException ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        }
    }


    void poiWriteProtectTest() throws IOException, GeneralSecurityException, InvalidFormatException {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream(); POIFSFileSystem fs = new POIFSFileSystem()) {
            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                XSSFSheet sheet = workbook.createSheet();
                XSSFRow row = sheet.createRow(0);
                XSSFCell cell = row.createCell(0);
                cell.setCellValue(testContent);
                workbook.write(bos);
                bos.flush();
                bos.close();
            }
            byte[] bytes = bos.toByteArray();
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            // EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);
            Encryptor enc = info.getEncryptor();
            enc.confirmPassword(secretKey);
            // Read in an existing OOXML file and write to encrypted output stream
            // don't forget to close the output stream otherwise the padding bytes aren't added
            try (OPCPackage opc = OPCPackage.open(new ByteArrayInputStream(bytes)); OutputStream os = enc.getDataStream(fs)) {
                opc.save(os);
            }
            // Write out the encrypted version
            try (FileOutputStream fos = new FileOutputStream(testFile)) {
                fs.writeFilesystem(fos);
            }
        }

    }

    void poiReadProtectTest() throws IOException {
        try (POIFSFileSystem fileSystem = new POIFSFileSystem(testFile)) {
            EncryptionInfo info = new EncryptionInfo(fileSystem);
            Decryptor d = Decryptor.getInstance(info);
            if (!d.verifyPassword(secretKey)) {
                throw new RuntimeException("Unable to process: document is encrypted");
            }
            // parse dataStream
            try (InputStream dataStream = d.getDataStream(fileSystem); XSSFWorkbook workbook = new XSSFWorkbook(dataStream)) {
                XSSFSheet sheet = workbook.getSheetAt(0);
                for (Row row : sheet) {
                    row.forEach(a -> System.out.println(a.getStringCellValue()));
                }
            }
        } catch (GeneralSecurityException ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        }
    }

    @AfterAll
    static void cleanup() {
        testFile.delete();
    }


    @Test
    public void fastexcelWrite_fastexcelRead() throws Exception {
        fastexcelWriteProtectTest();
        fastexcelReadProtectTest();
    }

    @Test
    public void fastexcelWrite_poiRead() throws Exception {
        fastexcelWriteProtectTest();
        poiReadProtectTest();
    }

    @Test
    public void poiWrite_fastexcelRead() throws Exception {
        poiWriteProtectTest();
        fastexcelReadProtectTest();
    }

    @Test
    public void poiWrite_poiRead() throws Exception {
        poiWriteProtectTest();
        poiReadProtectTest();
    }

}
