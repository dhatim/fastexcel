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
import org.junit.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.function.Consumer;

import static org.assertj.core.api.Assertions.assertThat;

public class Correctness {

    static byte[] writeWorkbook(Consumer<Workbook> consumer) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        Workbook wb = new Workbook(os, "Test", "1.0");
        consumer.accept(wb);
        wb.finish();
        return os.toByteArray();
    }

    @Test
    public void colToName() {
        assertThat(Range.colToString(26)).isEqualTo("AA");
        assertThat(Range.colToString(702)).isEqualTo("AAA");
        assertThat(Range.colToString(Worksheet.MAX_COLS - 1)).isEqualTo("XFD");
    }

    @Test(expected = IllegalArgumentException.class)
    public void noWorksheet() throws Exception {
        writeWorkbook(wb -> {
        });
    }

    @Test(expected = IllegalArgumentException.class)
    public void badVersion() {
        new Workbook(new NullOutputStream(), "Test", "1.0.1");
    }

    @Test
    public void singleEmptyWorksheet() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1"));
    }

    @Test
    public void worksheetWithNameLongerThan31Chars() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("01234567890123456789012345678901");
            assertThat(ws.getName()).isEqualTo("0123456789012345678901234567890");
        });
    }

    @Test
    public void worksheetsWithSameNames() throws Exception {
        byte[] data = writeWorkbook(wb -> {
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
    public void checkMaxRows() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS - 1, 0, "test"));
    }

    @Test
    public void checkMaxCols() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS - 1, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void exceedMaxRows() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(Worksheet.MAX_ROWS, 0, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void negativeRow() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(-1, 0, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void exceedMaxCols() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, Worksheet.MAX_COLS, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void negativeCol() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, -1, "test"));
    }

    @Test(expected = IllegalArgumentException.class)
    public void notSupportedTypeCell() throws Exception {
        writeWorkbook(wb -> wb.newWorksheet("Worksheet 1").value(0, 0, new Object()));
    }

    @Test(expected = IllegalArgumentException.class)
    public void invalidRange() throws Exception {
        writeWorkbook(wb -> {
            Worksheet ws = wb.newWorksheet("Worksheet 1");
            ws.range(-1, -1, Worksheet.MAX_COLS, Worksheet.MAX_ROWS);
        });
    }

    @Test
    public void reorderedRange() throws Exception {
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
    public void mergedRanges() throws Exception {
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

}
