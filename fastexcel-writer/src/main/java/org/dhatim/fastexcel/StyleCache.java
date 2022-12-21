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
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.function.Function;

/**
 * Thread-safe cache for shared styles. Check out
 * http://officeopenxml.com/SSstyles.php for more information about styles.
 */
final class StyleCache {

    private final ConcurrentMap<String, Integer> valueFormattings = new ConcurrentHashMap<>();
    private final ConcurrentMap<Font, Integer> fonts = new ConcurrentHashMap<>();
    private final ConcurrentMap<Fill, Integer> fills = new ConcurrentHashMap<>();
    private final ConcurrentMap<Border, Integer> borders = new ConcurrentHashMap<>();
    private final ConcurrentMap<Style, Integer> styles = new ConcurrentHashMap<>();
    private final ConcurrentMap<DifferentialFormat, Integer> dxfs = new ConcurrentHashMap<>();

    /**
     * Default constructor. Pre-cache Excel-reserved stuff.
     */
    StyleCache() {
        mergeAndCacheStyle(0, null, Font.DEFAULT, Fill.NONE, Border.NONE, null, null);
        cacheFill(Fill.GRAY125);
    }

    /**
     * Generic caching method.
     *
     * @param <T> Type of the cached objects.
     * @param cache Cache instance.
     * @param t Object being cached.
     * @param indexFunction Function to compute the index of the newly cached
     * object.
     * @return Index of the cached object.
     */
    private static <T> int cacheStuff(ConcurrentMap<T, Integer> cache, T t, Function<T, Integer> indexFunction) {
        return cache.computeIfAbsent(t, indexFunction);
    }

    /**
     * Caching method returning zero-based indexes.
     *
     * @param <T> Type of the cached objects.
     * @param cache Cache instance.
     * @param t Object being cached.
     * @return Index of the cached object.
     */
    private static <T> int cacheStuff(ConcurrentMap<T, Integer> cache, T t) {
        return cacheStuff(cache, t, k -> cache.size());
    }

    /**
     * Cache the given value formatting.
     *
     * @param s Value formatting.
     * @return Index of the cached format.
     */
    int cacheValueFormatting(String s) {
        if (s == null) {
            return 0;
        }
        return cacheStuff(valueFormattings, s, k -> valueFormattings.size() + 165);
    }

    /**
     * Cache the given font.
     *
     * @param f Font.
     * @return Index of the cached font.
     */
    int cacheFont(Font f) {
        return cacheStuff(fonts, f);
    }

    /**
     * Cache the given fill pattern.
     *
     * @param f Fill pattern.
     * @return Index of the cached fill pattern.
     */
    int cacheFill(Fill f) {
        return cacheStuff(fills, f);
    }

    /**
     * Cache the given border.
     *
     * @param b Border.
     * @return Index of the cached border.
     */
    int cacheBorder(Border b) {
        return cacheStuff(borders, b);
    }

    /**
     * Cache the given fill pattern as a shading color for rows.
     *
     * @param f Fill pattern.
     * @return Index of the cached fill pattern.
     */
    int cacheDxf(DifferentialFormat f) {
        return cacheStuff(dxfs, f);
    }

    int mergeAndCacheStyle(int currentStyle, String numberingFormat, Font font, Fill fill, Border border, Alignment alignment, Protection protection) {
        Style original = styles.entrySet().stream().filter(e -> e.getValue().equals(currentStyle)).map(Entry::getKey).findFirst().orElse(null);
        Style s = new Style(original, cacheValueFormatting(numberingFormat), cacheFont(font), cacheFill(fill), cacheBorder(border), alignment, protection);
        return cacheStuff(styles, s);
    }

    void replaceDefaultFont(Font font) {
        fonts.entrySet().removeIf(entry->entry.getValue()==0);
        fonts.putIfAbsent(font,0);
    }

    /**
     * Write a cache as an XML element.
     *
     * @param <T> Type of the cached objects.
     * @param w Output writer.
     * @param cache Cache instance.
     * @param name Name of the XML element.
     * @param consumer Consumer to write cached elements.
     * @throws IOException If an I/O error occurs.
     */
    private static <T> void writeCache(Writer w, Map<T, Integer> cache, String name, ThrowingConsumer<Entry<T, Integer>> consumer) throws IOException {
        w.append('<').append(name).append(" count=\"").append(cache.size()).append("\">");
        List<Entry<T, Integer>> entries = new ArrayList<>(cache.entrySet());
        entries.sort(Comparator.comparingInt(Entry::getValue));
        for (Entry<T, Integer> e : entries) {
            consumer.accept(e);
        }
        w.append("</").append(name).append('>');
    }

    /**
     * Write this style cache as an XML file.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        writeCache(w, valueFormattings, "numFmts", e -> w.append("<numFmt numFmtId=\"").append(e.getValue()).append("\" formatCode=\"").append(e.getKey()).append("\"/>"));
        writeCache(w, fonts, "fonts", e -> e.getKey().write(w));
        writeCache(w, fills, "fills", e -> e.getKey().write(w));
        writeCache(w, borders, "borders", e -> e.getKey().write(w));
        w.append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
        writeCache(w, styles, "cellXfs", e -> e.getKey().write(w));
        writeCache(w, dxfs, "dxfs", e -> {
            e.getKey().write(w);
        });
        w.append("</styleSheet>");
    }
}
