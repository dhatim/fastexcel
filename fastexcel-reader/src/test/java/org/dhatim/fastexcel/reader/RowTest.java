package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

import static org.assertj.core.api.Assertions.assertThat;
import static org.dhatim.fastexcel.reader.Resources.open;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

public class RowTest {

	@Test
	public void testGetCellByAddress() throws IOException {
		try (InputStream inputStream = open("/xlsx/issue.xlsx");
		     ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
			Sheet firstSheet = excel.getFirstSheet();
			Row row = firstSheet.read().get(0);
			assertNotNull(row);
			CellAddress addressA1 = new CellAddress("A1");
			Cell cellIndex0 = row.getCell(0);
			Cell cellA1 = row.getCell(addressA1);
			assertEquals(addressA1, cellIndex0.getAddress());
			assertEquals(cellIndex0, cellA1);
		}
	}

    @ParameterizedTest
    @CsvSource({
            "0, false, 1, Lorem, 43101, 1+A1, true",
            "1, true, 2, ipsum, 43102, 1+A2, false",
            "2, true, 3, dolor, 43103, 1+A3, true",
            "3, false, 4, sit, 43104, 1+A4, false",
            "4, false, 5, amet, 43105, 1+A5, true",
            "5, false, 6, consectetur, 43106, 1+A6, false",
            "6, true, 7, adipiscing, 43107, 1+A7, true",
            "7, true, 8, elit, 43108, 1+A8, false",
            "8, true, 9, Ut, 43109, 1+A9, true",
            "9, false, 10, nec, 43110, 1+A10, false"
    })
    public void shouldGetVisibleAndHiddenRows(int index, boolean isHidden, String cellIndex0, String cellIndex1, int cellIndex2, String cellIndex3,
            boolean cellIndex4) throws IOException {
        try (InputStream inputStream = open("/xlsx/simple-with-hidden-rows.xlsx");
                ReadableWorkbook excel = new ReadableWorkbook(inputStream)) {
            Sheet firstSheet = excel.getFirstSheet();

            Row row = firstSheet.read().get(index);

            assertThat(row.isHidden()).isEqualTo(isHidden);
            assertThat(row.getCell(0)).satisfies(cell -> {
                assertThat(cell.getType()).isEqualTo(CellType.NUMBER);
                assertThat(cell.asNumber()).isEqualTo(new BigDecimal(cellIndex0));
            });
            assertThat(row.getCell(1)).satisfies(cell -> {
                assertThat(cell.getType()).isEqualTo(CellType.STRING);
                assertThat(cell.asString()).isEqualTo(cellIndex1);
            });
            assertThat(row.getCell(2)).satisfies(cell -> {
                assertThat(cell.getType()).isEqualTo(CellType.NUMBER);
                assertThat(cell.asNumber()).isEqualTo(new BigDecimal(cellIndex2));
            });
            assertThat(row.getCell(3)).satisfies(cell -> {
                assertThat(cell.getType()).isEqualTo(CellType.FORMULA);
                assertThat(cell.getFormula()).isEqualTo(cellIndex3);
            });
            assertThat(row.getCell(4)).satisfies(cell -> {
                assertThat(cell.getType()).isEqualTo(CellType.BOOLEAN);
                assertThat(cell.asBoolean()).isEqualTo(cellIndex4);
            });
        }
    }
}
