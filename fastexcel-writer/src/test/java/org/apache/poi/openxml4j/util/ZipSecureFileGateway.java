package org.apache.poi.openxml4j.util;

public class ZipSecureFileGateway {
    /**
     * Suppresses the following exception:
     * "java.io.IOException: Zip bomb detected! The file would exceed the max size of the expanded data in the zip-file.
     * This may indicates that the file is used to inflate memory usage and thus could pose a security risk.
     * You can adjust this limit via ZipSecureFile.setMaxEntrySize() if you need to work with files which are very large.
     * Uncompressed size: 4294970702, Raw/compressed size: 492446720
     * Limits: MAX_ENTRY_SIZE: 4294967295, Entry: xl/worksheets/sheet1.xml"
     */
    public static void disableZipBombDetection() {
        ZipSecureFile.MAX_ENTRY_SIZE = Long.MAX_VALUE;
    }
}
