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

import java.io.IOException;
import java.util.Date;
import java.util.zip.Deflater;

import org.apache.commons.io.output.CountingOutputStream;
import org.apache.commons.io.output.NullOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openjdk.jmh.annotations.Benchmark;

/**
 * Compare write performance between this library and
 * <a href="https://poi.apache.org/">Apache POI</a>.
 */
public class Benchmarks {

    private static final int NB_ROWS = 100_000;

    @Benchmark
    public Object poiNoStreaming() throws Exception {
        return poiPopulate(new XSSFWorkbook());
    }

    @Benchmark
    public Object poiStreaming() throws Exception {
        return poiPopulate(new SXSSFWorkbook());
    }

    @Benchmark
    public Object fastExcel() throws Exception {
        return fastExcel(Deflater.DEFAULT_COMPRESSION);
    }

    @Benchmark
    public Object fastExcelFastCompression() throws Exception {
        return fastExcel(Deflater.BEST_SPEED);
    }

    @Benchmark
    public Object fastExcelCompression4() throws Exception {
        return fastExcel(4);
    }


    @Test
    public void pickCompressionLevel() throws IOException {
        for (int i = 0; i < 5; i++) {
            int[] results = new int[Deflater.BEST_COMPRESSION];
            for (int level = 0; level < results.length; level++) {
                long start = System.currentTimeMillis();
                int size = fastExcel(level);
                long end = System.currentTimeMillis();
                System.out.println(level + "; " + size + "; " + (end - start));
            }
            long start = System.currentTimeMillis();
            int size = fastExcel(Deflater.DEFAULT_COMPRESSION);
            long end = System.currentTimeMillis();
            System.out.println(Deflater.DEFAULT_COMPRESSION+ "; " + size + "; " + (end - start));
            System.out.println();
        }
    }

    private int fastExcel(int compression) throws IOException {
        CountingOutputStream count = new CountingOutputStream(new NullOutputStream());
        Workbook wb = new Workbook(count, "Perf", "1.0");
        wb.setCompressionLevel(compression);
        Worksheet ws = wb.newWorksheet("Sheet 1");
        for (int r = 0; r < NB_ROWS; ++r) {
            ws.value(r, 0, r);
            ws.value(r, 1, Integer.toString(r % 1000));
            ws.value(r, 2, r / 87.0);
            ws.value(r, 3, new Date(1549915044));
        }
        ws.range(0, 3, NB_ROWS - 1, 3).style().format("yyyy-mm-dd hh:mm:ss").set();
        wb.finish();
        return count.getCount();
    }

    private int poiPopulate(org.apache.poi.ss.usermodel.Workbook wb) throws Exception {
        Sheet ws = wb.createSheet("Sheet 1");
        CellStyle dateStyle = wb.createCellStyle();
        dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        for (int r = 0; r < NB_ROWS; ++r) {
            Row row = ws.createRow(r);
            row.createCell(0).setCellValue(r);
            row.createCell(1).setCellValue(Integer.toString(r % 1000));
            row.createCell(2).setCellValue(r / 87.0);
            Cell c = row.createCell(3);
            c.setCellStyle(dateStyle);
            c.setCellValue(new Date(1549915044));
        }
        CountingOutputStream count = new CountingOutputStream(new NullOutputStream());
        wb.write(count);
        return count.getCount();
    }
}
