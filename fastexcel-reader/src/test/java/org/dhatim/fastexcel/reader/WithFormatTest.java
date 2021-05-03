package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Optional;

import static org.dhatim.fastexcel.reader.Resources.open;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class WithFormatTest {

  @Test
  void testFile() throws IOException {
    try (InputStream inputStream = open("/xlsx/withStyle.xlsx");
         ReadableWorkbook excel = new ReadableWorkbook(inputStream, new ReadingOptions(true, false))) {
      Optional<Sheet> sheet = excel.getActiveSheet();
      assertTrue(sheet.isPresent());
      Iterator<Row> it = sheet.get().openStream().iterator();
      assertTrue(it.hasNext());
      it.next();
      assertTrue(it.hasNext());
      Row row = it.next();
      Iterator<Cell> itCell = row.stream().iterator();
      assertTrue(itCell.hasNext());
      Cell cell = itCell.next();
      assertEquals(cell.getDataFormatId(), (Integer) 164);
      assertEquals(cell.getDataFormatString(), "General");
      assertTrue(itCell.hasNext());
      cell = itCell.next();
      assertEquals(cell.getDataFormatId(), (Integer) 166);
      assertEquals(cell.getDataFormatString(), "DD\\-MM\\-YYYY");
    }
  }
}
