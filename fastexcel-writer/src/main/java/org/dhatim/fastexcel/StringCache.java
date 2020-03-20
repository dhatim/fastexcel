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
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map.Entry;
import java.util.stream.Stream;

import static java.util.Comparator.comparingInt;

/**
 * Thread-safe cache for shared strings. Each string is uniquely identified by
 * an integer. See {@link CachedString}.
 */
class StringCache {

    /**
     * Number of strings, including duplicates.
     */
    private long count;
    /**
     * Map giving string index for each unique string.
     */
    private final HashMap<String, CachedString> strings = new HashMap<>();

    /**
     * Add a string to this cache.
     *
     * @param s String to cache.
     * @return Cached string.
     */
    CachedString cacheString(String s) {
        CachedString result;
        synchronized (strings) {
            ++count;
            result = strings.get(s);
            if (result == null) {
                result = new CachedString(s, strings.size());
                strings.put(s, result);
            }
        }
        return result;
    }

    /**
     * Write this cache as an XML file.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"").append(count).append("\" uniqueCount=\"").append(strings.size()).append("\">");
        Stream<String> sortedStrings = strings.entrySet().stream()
                .sorted(comparingInt(e -> e.getValue().getIndex()))
                .map(Entry::getKey);
        Iterator<String> it = sortedStrings.iterator();
        while (it.hasNext()) {
            w.append("<si><t>").appendEscaped(it.next()).append("</t></si>");
        }
        w.append("</sst>");
    }
}
