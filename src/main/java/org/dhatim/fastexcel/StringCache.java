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
import java.util.HashMap;
import java.util.Map;

/**
 * Thread-safe cache for shared strings. Each string is uniquely identified by
 * an integer.
 */
class StringCache {

    /**
     * Number of strings, including duplicates.
     */
    private long count = 0;
    /**
     * Map giving string indexes for each unique string.
     */
    private final Map<String, Integer> stringToInt = new HashMap<>();
    /**
     * Map giving strings for each unique string index.
     */
    private final Map<Integer, String> intToString = new HashMap<>();

    /**
     * Add a string to this cache.
     *
     * @param s String to cache.
     * @return Index of cached string.
     */
    int cacheString(String s) {
        synchronized (this) {
            ++count;
            Integer index = stringToInt.get(s);
            if (index == null) {
                index = stringToInt.size();
                stringToInt.put(s, index);
                intToString.put(index, s);
            }
            return index;
        }
    }

    /**
     * Convert string index into string.
     *
     * @param index Index value.
     * @return Cached string or {@code null} if not found.
     */
    String getString(int index) {
        synchronized (this) {
            return intToString.get(index);
        }
    }

    /**
     * Write this cache as an XML file.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"").append(count).append("\" uniqueCount=\"").append(stringToInt.size()).append("\">");
        for (int i = 0; i < stringToInt.size(); ++i) {
            w.append("<si><t>").appendEscaped(intToString.get(i)).append("</t></si>");
        }
        w.append("</sst>");
    }
}
