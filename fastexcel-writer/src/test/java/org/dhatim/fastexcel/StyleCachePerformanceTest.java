package org.dhatim.fastexcel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

class StyleCachePerformanceTest {

    @Test
    void styleCacheScalesWellWithManyStyles() throws IOException {
        int numStyles = 1000;
        int mergesPerStyle = 10;
        
        long startTime = System.nanoTime();
        
        try (ByteArrayOutputStream out = new ByteArrayOutputStream();
             Workbook wb = new Workbook(out, "PerfTest", "1.0")) {
            
            Worksheet ws = wb.newWorksheet("Sheet 1");
            
            for (int styleIdx = 0; styleIdx < numStyles; styleIdx++) {
                String fontName = "Font" + (styleIdx % 20);
                int fontSize = 10 + (styleIdx % 5);
                String fillColor = String.format("#%06X", (styleIdx * 1000) % 0xFFFFFF);
                
                for (int merge = 0; merge < mergesPerStyle; merge++) {
                    StyleSetter styleSetter = ws.range(styleIdx % 100, (styleIdx * 2) % 50, 
                            (styleIdx % 100) + 1, ((styleIdx * 2) % 50) + 1)
                      .style()
                      .fontName(fontName)
                      .fontSize(fontSize)
                      .fillColor(fillColor);
                    if (styleIdx % 2 == 0) {
                        styleSetter.bold();
                    }
                    if (styleIdx % 3 == 0) {
                        styleSetter.italic();
                    }
                    styleSetter.horizontalAlignment(styleIdx % 2 == 0 ? "center" : "left").set();
                }
            }
            
            for (int i = 0; i < 100; i++) {
                ws.value(i, 0, "Data " + i);
            }
        }
        
        long duration = System.nanoTime() - startTime;
        long durationMs = duration / 1_000_000;
        
        assertThat(durationMs).isLessThan(5000);
        
        System.out.println(String.format(
            "Created %d styles with %d merges each in %d ms (%.2f ms per style merge)",
            numStyles, mergesPerStyle, durationMs, 
            (double) durationMs / (numStyles * mergesPerStyle)
        ));
    }

    @Test
    void styleCacheHandlesManyUniqueStyles() throws IOException {
        int numStyles = 2000;
        
        long startTime = System.nanoTime();
        
        try (ByteArrayOutputStream out = new ByteArrayOutputStream();
             Workbook wb = new Workbook(out, "PerfTest", "1.0")) {
            
            Worksheet ws = wb.newWorksheet("Sheet 1");
            
            for (int styleIdx = 0; styleIdx < numStyles; styleIdx++) {
                ws.range(styleIdx % 100, styleIdx % 10, 
                        (styleIdx % 100) + 1, (styleIdx % 10) + 1)
                  .style()
                  .fontName("Font" + styleIdx)
                  .fontSize(10 + (styleIdx % 20))
                  .fillColor(String.format("#%06X", styleIdx % 0xFFFFFF))
                  .set();
            }
        }
        
        long duration = System.nanoTime() - startTime;
        long durationMs = duration / 1_000_000;
        
        assertThat(durationMs).isLessThan(3000);
        
        System.out.println(String.format(
            "Created %d unique styles in %d ms (%.2f ms per style)",
            numStyles, durationMs, (double) durationMs / numStyles
        ));
    }
}
