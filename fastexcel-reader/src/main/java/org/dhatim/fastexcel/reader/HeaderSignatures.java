package org.dhatim.fastexcel.reader;

import java.util.Arrays;

class HeaderSignatures {
  static final byte[] OLE_2_SIGNATURE = new byte[]{-48, -49, 17, -32, -95, -79, 26, -31};
  static final byte[] OOXML_FILE_HEADER = new byte[]{0x50, 0x4b, 0x03, 0x04};

  static boolean isHeader(byte[] bytes, byte[] header) {
    requireLength(bytes, header.length);
    return Arrays.equals(
        Arrays.copyOf(bytes, header.length),
        header
    );
  }
  private static void requireLength(byte[] bytes, int requiredLength) {
    if (bytes.length < requiredLength) {
      throw new IllegalArgumentException("Insufficient header bytes");
    }
  }
}
