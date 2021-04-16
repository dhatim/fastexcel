package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import static org.dhatim.fastexcel.reader.Resources.open;
import static org.junit.jupiter.api.Assertions.*;

public class SheetInfoTest {

  @Test
  void testWithHiddenSheet() throws IOException {
    try (InputStream inputStream = open("/xlsx/withHidden.xlsx");
         ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
      assertTrue(excel.getActiveSheet().isPresent());
      assertEquals(excel.getActiveSheet().get().getName(), "ActiveSheet");
      assertEquals(excel.getSheets().count(), 4L);
      Iterator<Sheet> it = excel.getSheets().iterator();
      assertTrue(it.hasNext());
      assertEquals(it.next().getVisibility(), SheetVisibility.VISIBLE);
      assertTrue(it.hasNext());
      assertEquals(it.next().getVisibility(), SheetVisibility.HIDDEN);
      assertTrue(it.hasNext());
      assertEquals(it.next().getVisibility(), SheetVisibility.VISIBLE);
      assertTrue(it.hasNext());
      assertEquals(it.next().getVisibility(), SheetVisibility.VISIBLE);
      assertFalse(it.hasNext());
    }
  }

  @Test
  void testWithVeryHidden() throws IOException {
    try (InputStream inputStream = open("/xlsx/VeryHidden.xlsx");
         ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
      assertFalse(excel.getActiveSheet().isPresent());
      assertEquals(excel.getSheets().count(), 2L);
      Iterator<Sheet> it = excel.getSheets().iterator();
      assertTrue(it.hasNext());
      assertEquals(it.next().getVisibility(), SheetVisibility.VISIBLE);
      assertTrue(it.hasNext());
      assertEquals(it.next().getVisibility(), SheetVisibility.VERY_HIDDEN);
      assertFalse(it.hasNext());
    }
  }
}
