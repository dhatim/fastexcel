package org.dhatim.fastexcel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Map.Entry;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

import static org.assertj.core.api.Assertions.assertThat;

class StyleCacheBeforeAfterTest {

    private static class OldStyleCache {
        private final ConcurrentMap<Style, Integer> styles = new ConcurrentHashMap<>();
        
        int mergeAndCacheStyleOldWay(int currentStyle, String numberingFormat, Font font, 
                                     Fill fill, Border border, Alignment alignment, Protection protection,
                                     StyleCacheHelper helper) {
            Style original = styles.entrySet().stream()
                    .filter(e -> e.getValue().equals(currentStyle))
                    .map(Entry::getKey)
                    .findFirst()
                    .orElse(null);
            
            Style s = new Style(original, 
                    helper.cacheValueFormatting(numberingFormat),
                    helper.cacheFont(font),
                    helper.cacheFill(fill),
                    helper.cacheBorder(border),
                    alignment, protection);
            
            return styles.computeIfAbsent(s, k -> styles.size());
        }
    }

    private static class NewStyleCache {
        private final ConcurrentMap<Style, Integer> styles = new ConcurrentHashMap<>();
        private final ConcurrentMap<Integer, Style> styleIndexToStyle = new ConcurrentHashMap<>();
        
        int mergeAndCacheStyleNewWay(int currentStyle, String numberingFormat, Font font,
                                    Fill fill, Border border, Alignment alignment, Protection protection,
                                    StyleCacheHelper helper) {
            Style original = styleIndexToStyle.get(currentStyle);
            
            Style s = new Style(original,
                    helper.cacheValueFormatting(numberingFormat),
                    helper.cacheFont(font),
                    helper.cacheFill(fill),
                    helper.cacheBorder(border),
                    alignment, protection);
            
            Integer index = styles.computeIfAbsent(s, k -> styles.size());
            styleIndexToStyle.putIfAbsent(index, s);
            return index;
        }
    }

    private interface StyleCacheHelper {
        int cacheValueFormatting(String s);
        int cacheFont(Font f);
        int cacheFill(Fill f);
        int cacheBorder(Border b);
    }

    @Test
    void compareOldVsNewImplementation() {
        int numStyles = 1000;
        int mergesPerStyle = 10;
        
        StyleCacheHelper helper = new StyleCacheHelper() {
            @Override
            public int cacheValueFormatting(String s) { return s != null ? 1 : 0; }
            @Override
            public int cacheFont(Font f) { return 0; }
            @Override
            public int cacheFill(Fill f) { return 0; }
            @Override
            public int cacheBorder(Border b) { return 0; }
        };
        
        OldStyleCache oldCache = new OldStyleCache();
        NewStyleCache newCache = new NewStyleCache();
        
        for (int i = 0; i < 100; i++) {
            oldCache.mergeAndCacheStyleOldWay(0, null, Font.DEFAULT, Fill.NONE, Border.NONE, null, null, helper);
            newCache.mergeAndCacheStyleNewWay(0, null, Font.DEFAULT, Fill.NONE, Border.NONE, null, null, helper);
        }
        
        int currentStyle = 0;
        for (int i = 0; i < numStyles; i++) {
            currentStyle = oldCache.mergeAndCacheStyleOldWay(
                    currentStyle, "fmt" + i, Font.DEFAULT, Fill.NONE, Border.NONE, null, null, helper);
            newCache.mergeAndCacheStyleNewWay(
                    currentStyle, "fmt" + i, Font.DEFAULT, Fill.NONE, Border.NONE, null, null, helper);
        }
        
        long oldStart = System.nanoTime();
        currentStyle = 0;
        for (int styleIdx = 0; styleIdx < numStyles; styleIdx++) {
            for (int merge = 0; merge < mergesPerStyle; merge++) {
                currentStyle = oldCache.mergeAndCacheStyleOldWay(
                        currentStyle,
                        "format" + (styleIdx % 10),
                        Font.DEFAULT,
                        Fill.NONE,
                        Border.NONE,
                        null, null, helper);
            }
        }
        long oldDuration = System.nanoTime() - oldStart;
        
        long newStart = System.nanoTime();
        currentStyle = 0;
        for (int styleIdx = 0; styleIdx < numStyles; styleIdx++) {
            for (int merge = 0; merge < mergesPerStyle; merge++) {
                currentStyle = newCache.mergeAndCacheStyleNewWay(
                        currentStyle,
                        "format" + (styleIdx % 10),
                        Font.DEFAULT,
                        Fill.NONE,
                        Border.NONE,
                        null, null, helper);
            }
        }
        long newDuration = System.nanoTime() - newStart;
        
        System.out.println("\n=== Style Cache Performance Comparison ===");
        System.out.printf("Styles in cache: %d%n", numStyles);
        System.out.printf("Merges per style: %d%n", mergesPerStyle);
        System.out.printf("Total operations: %d%n", numStyles * mergesPerStyle);
        System.out.println("\nOLD (O(n) linear search):");
        System.out.printf("  Time: %d ms (%.2f ns per operation)%n",
                oldDuration / 1_000_000, (double) oldDuration / (numStyles * mergesPerStyle));
        System.out.println("\nNEW (O(1) hash map lookup):");
        System.out.printf("  Time: %d ms (%.2f ns per operation)%n",
                newDuration / 1_000_000, (double) newDuration / (numStyles * mergesPerStyle));
        
        double speedup = (double) oldDuration / newDuration;
        System.out.printf("\nSpeedup: %.2fx faster%n", speedup);
        
        assertThat(newDuration).isLessThanOrEqualTo(oldDuration);
        assertThat(speedup).isGreaterThan(1.5);
    }

    @Test
    void demonstrateRealWorldScenario() throws IOException {
        int numUniqueStyles = 2000;
        
        System.out.println("\n=== Real-World Scenario: Many Unique Styles ===");
        System.out.printf("Creating workbook with %d unique styles...%n", numUniqueStyles);
        
        long startTime = System.nanoTime();
        
        try (ByteArrayOutputStream out = new ByteArrayOutputStream();
             Workbook wb = new Workbook(out, "PerfTest", "1.0")) {
            
            Worksheet ws = wb.newWorksheet("Sheet 1");
            
            for (int styleIdx = 0; styleIdx < numUniqueStyles; styleIdx++) {
                StyleSetter styleSetter = ws.range(styleIdx % 100, styleIdx % 10, 
                        (styleIdx % 100) + 1, (styleIdx % 10) + 1)
                  .style()
                  .fontName("Font" + (styleIdx % 50))
                  .fontSize(10 + (styleIdx % 20))
                  .fillColor(String.format("#%06X", styleIdx % 0xFFFFFF));
                if (styleIdx % 3 == 0) {
                    styleSetter.bold();
                }
                if (styleIdx % 4 == 0) {
                    styleSetter.italic();
                }
                styleSetter.horizontalAlignment(styleIdx % 2 == 0 ? "center" : "left").set();
            }
            
            for (int i = 0; i < 100; i++) {
                ws.value(i, 0, "Data " + i);
            }
        }
        
        long duration = System.nanoTime() - startTime;
        long durationMs = duration / 1_000_000;
        
        System.out.printf("Completed in %d ms%n", durationMs);
        System.out.printf("Average: %.2f ms per style%n", (double) durationMs / numUniqueStyles);
        
        assertThat(durationMs).isLessThan(5000);
    }
}
