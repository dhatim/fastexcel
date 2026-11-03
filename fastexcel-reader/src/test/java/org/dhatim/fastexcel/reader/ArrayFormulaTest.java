package org.dhatim.fastexcel.reader;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTimeoutPreemptively;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

public class ArrayFormulaTest {

	@Test
	void spilledCellsGetTheFormulaOfTheirRange() throws Exception {
		List<Row> rows = read(
				"<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:B2\">FILTER(X:X,Y:Y=1)</f><v>1</v></c><c r=\"B1\"><v>2</v></c><c r=\"C1\"><v>3</v></c></row>"
						+ "<row r=\"2\"><c r=\"A2\"><v>4</v></c><c r=\"B2\"><v>5</v></c><c r=\"C2\"><v>6</v></c></row>");

		assertEquals("FILTER(X:X,Y:Y=1)", rows.get(0).getCell(1).getFormula());
		assertEquals("FILTER(X:X,Y:Y=1)", rows.get(1).getCell(0).getFormula());
		assertEquals("FILTER(X:X,Y:Y=1)", rows.get(1).getCell(1).getFormula());
		assertEquals(CellType.FORMULA, rows.get(1).getCell(1).getType());
		assertNull(rows.get(0).getCell(2).getFormula());
		assertNull(rows.get(1).getCell(2).getFormula());
	}

	@Test
	void rangeDeclaredBelowDoesNotApplyToCellsAlreadyRead() throws Exception {
		List<Row> rows = read(
				"<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
						+ "<row r=\"2\"><c r=\"A2\"><f t=\"array\" ref=\"A1:A3\">LATE()</f><v>2</v></c></row>"
						+ "<row r=\"3\"><c r=\"A3\"><v>3</v></c></row>");

		assertNull(rows.get(0).getCell(0).getFormula());
		assertEquals("LATE()", rows.get(2).getCell(0).getFormula());
	}

	@Test
	void rangeStillAppliesWhenRowsAreOutOfOrder() throws Exception {
		List<Row> rows = read(
				"<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A3\">EARLY()</f><v>1</v></c></row>"
						+ "<row r=\"5\"><c r=\"A5\"><v>5</v></c></row>"
						+ "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row>");

		assertNull(rows.get(1).getCell(0).getFormula());
		assertEquals("EARLY()", rows.get(2).getCell(0).getFormula());
	}

	@Test
	void manyArrayFormulasDoNotSlowDownReading() {
		// 2,000 rows x 60 columns, 15 single-cell array formulas per row (30,000 in total) before plain numbers.
		// Scanning every range for each cell took about 30 seconds here.
		StringBuilder sheetData = new StringBuilder();
		for (int r = 1; r <= 2000; r++) {
			sheetData.append("<row r=\"").append(r).append("\">");
			for (int c = 0; c < 60; c++) {
				String ref = CellAddress.convertNumToColString(c) + r;
				sheetData.append("<c r=\"").append(ref).append("\">");
				if (c < 15) {
					sheetData.append("<f t=\"array\" ref=\"").append(ref).append("\">INDEX(FILTER(X:X,Y:Y=").append(r).append("),1)</f>");
				}
				sheetData.append("<v>").append(r + c).append("</v></c>");
			}
			sheetData.append("</row>");
		}

		List<Row> rows = assertTimeoutPreemptively(Duration.ofSeconds(10), () -> read(sheetData.toString()));
		assertEquals(2000, rows.size());
	}

	private static List<Row> read(String sheetData) throws Exception {
		String xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
				+ sheetData + "</sheetData></worksheet>";
		RowSpliterator it = new RowSpliterator(null, new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));
		List<Row> rows = new ArrayList<>();
		while (it.tryAdvance(rows::add)) {
			// read all rows
		}
		return rows;
	}
}
