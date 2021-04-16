package org.dhatim.fastexcel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

public class WriteFileTest {

  @Test
  void testWrite() throws IOException {
    try (FileOutputStream os = new FileOutputStream("/home/mathieu/testFormulatWithoutValue.xlsx")) {
      Workbook wb = new Workbook(os, "Test", "1.0");
      wb.beginFile("testFormulaWithoutValue");
      Worksheet sheet = wb.newWorksheet("sheet");
      sheet.value(0, 0, new BigDecimal("1"));
      sheet.value(0, 1, new BigDecimal("2"));
      sheet.formula(0,2, "SUM(A1:B1)");
      wb.finish();
    }
  }
}
