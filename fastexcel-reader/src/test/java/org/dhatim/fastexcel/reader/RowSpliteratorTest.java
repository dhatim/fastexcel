package org.dhatim.fastexcel.reader;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.InputStream;

import org.junit.jupiter.api.Test;

public class RowSpliteratorTest {

	@Test
	void testBlankCells() throws Exception {
		InputStream is = Resources.open("/xml/blank_cells.xml");
		RowSpliterator it = new RowSpliterator(null, is);
		it.tryAdvance(row -> {
			assertEquals(8, row.getCellCount());
			Cell cell = row.getCell(7);
			assertEquals("H1", cell.getAddress().toString());
		});
		it.tryAdvance(row -> {
			assertEquals(8, row.getCellCount());
			Cell cell = row.getCell(7);
			assertEquals("H2", cell.getAddress().toString());
		});
	}
	
}

