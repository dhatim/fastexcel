/*
 * Copyright 2016 Dhatim.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.dhatim.fastexcel;

import org.apache.commons.io.output.NullOutputStream;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.util.Date;
import java.util.function.Consumer;
import java.util.stream.IntStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

class CorrectnessTest {

    static byte[] writeWorkbook(Consumer<Workbook> consumer) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        Workbook wb = new Workbook(os, "Test", "1.0");
        consumer.accept(wb);
        wb.finish();
        return os.toByteArray();
    }

    @Test
    void colToName() {
        assertThat(Range.colToString(26)).isEqualTo("AA");
        assertThat(Range.colToString(702)).isEqualTo("AAA");
        assertThat(Range.colToString(Worksheet.MAX_COLS - 1)).isEqualTo("XFD");
    }

    @Test
    void noWorksheet() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> {
            });
        });
    }

    @Test
    void badVersion() {
        assertThrows(IllegalArgumentException.class, () -> {
            new Workbook(new NullOutputStream(), "Test", "1.0.1");
        });
    }

    @Test
    void singleEmptyWorksheet() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1"));
    }

    @Test
    void worksheetWithNameLongerThan31Chars() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("01234567890123456789012345678901");
            assertThat(ws.getName()).isEqualTo("0123456789012345678901234567890");
        });
    }

    @Test
    void worksheetsWithSameNames() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("01234567890123456789012345678901");
            assertThat(ws.getName()).isEqualTo("0123456789012345678901234567890");
            ws = wb.newWorksheet("0123456789012345678901234567890");
            assertThat(ws.getName()).isEqualTo("01234567890123456789012345678_1");
            ws = wb.newWorksheet("01234567890123456789012345678_1");
            assertThat(ws.getName()).isEqualTo("01234567890123456789012345678_2");
            wb.newWorksheet("abc");
            ws = wb.newWorksheet("abc");
            assertThat(ws.getName()).isEqualTo("abc_1");
        });
    }

    @Test
    void checkMaxRows() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS - 1, 0, "test"));
    }

    @Test
    void checkMaxCols() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS - 1, "test"));
    }

    @Test
    void exceedMaxRows() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS, 0, "test"));
        });
    }

    @Test
    void negativeRow() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(-1, 0, "test"));
        });
    }

    @Test
    void exceedMaxCols() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS, "test"));
        });
    }

    @Test
    void negativeCol() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, -1, "test"));
        });
    }

    @Test
    void invalidRange() {
        assertThrows(IllegalArgumentException.class, () -> {
            writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.range(-1, -1, Worksheet.MAX_COLS, Worksheet.MAX_ROWS);
            });
        });
    }

    @Test
    void zoomTooSmall() {
        assertThrows(IllegalArgumentException.class, () -> {
            //if (scale >= 10 && scale <= 400) {
            writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.setZoom(9);
            });
        });
    }

    @Test
    void zoomTooBig() {
        assertThrows(IllegalArgumentException.class, () -> {
            //if (scale >= 10 && scale <= 400) {
            writeWorkbook(wb -> {
                Worksheet ws = wb.newWorksheet("Worksheet 1");
                ws.setZoom(401);
            });
        });
    }

    @Test
    void reorderedRange() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            int top = 0;
            int left = 1;
            int bottom = 10;
            int right = 11;
            Range range = ws.range(top, left, bottom, right);
            Range otherRange = ws.range(bottom, right, top, left);
            assertThat(range).isEqualTo(otherRange);
            assertThat(range.getTop()).isEqualTo(top);
            assertThat(range.getLeft()).isEqualTo(left);
            assertThat(range.getBottom()).isEqualTo(bottom);
            assertThat(range.getRight()).isEqualTo(right);
            assertThat(otherRange.getTop()).isEqualTo(top);
            assertThat(otherRange.getLeft()).isEqualTo(left);
            assertThat(otherRange.getBottom()).isEqualTo(bottom);
            assertThat(otherRange.getRight()).isEqualTo(right);
        });
    }

    @Test
    void mergedRanges() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.value(0, 0, "One");
            ws.value(0, 1, "Two");
            ws.value(0, 2, "Three");
            ws.value(1, 0, "Merged");
            ws.range(1, 0, 1, 2).style().merge().set();
            ws.range(1, 0, 1, 2).merge();
            ws.style(1, 0).horizontalAlignment("center").set();
        });
    }

    @Test
    void shouldBeAbleToNullifyCell() throws IOException {
        writeWorkbook(wb ->{
            Worksheet ws = wb.newWorksheet("Sheet 1");
            ws.value(0,0, "One");
            ws.value(1,0, 42);
            ws.value(2,0, true);
            ws.value(3,0, new Date());
            ws.value(4,0, LocalDate.now());
            ws.value(5,0, LocalDateTime.now());
            ws.value(6,0, ZonedDateTime.now());
            for (int r = 0; r <= 6; r++) {
              assertThat(ws.cell(r, 0).getValue()).isNotNull();
            }
            ws.value(0,0, (Boolean) null);
            ws.value(1,0, (Number) null);
            ws.value(2,0, (String) null);
            ws.value(3,0, (LocalDate) null);
            ws.value(4,0, (ZonedDateTime) null);
            ws.value(5,0, (LocalDateTime) null);
            ws.value(6,0, (LocalDate) null);
            for (int r = 0; r <= 6; r++) {
              assertThat(ws.cell(r, 0).getValue()).isNull();
            }
        });
    }
}
