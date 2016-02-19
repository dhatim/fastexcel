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

/**
 * A cell contains a value and a cached style index.
 */
class Cell {

    /**
     * Cell value.
     */
    Object value;

    /**
     * Cached style index.
     */
    int style = 0;

    /**
     * Write this cell as an XML element.
     *
     * @param w Output writer.
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w, int r, int c) throws IOException {
        if (value != null || style != 0) {
            w.append("<c r=\"").append(Range.colToString(c)).append(r + 1).append('\"');
            if (style != 0) {
                w.append(" s=\"").append(style).append('\"');
            }
            if (value != null && !(value instanceof Formula)) {
                w.append(" t=\"").append((value instanceof String || value instanceof CachedString) ? 's' : 'n').append('\"');
            }
            w.append(">");
            if (value instanceof Formula) {
                w.append("<f>").append(((Formula) value).getExpression()).append("</f>");
            } else {
                w.append("<v>");
                if (value instanceof String) {
                } else if (value instanceof CachedString) {
                    w.append(((CachedString) value).getIndex());
                } else if (value instanceof Integer) {
                    w.append((int) value);
                } else if (value instanceof Long) {
                    w.append((long) value);
                } else if (value instanceof Double) {
                    w.append((double) value);
                } else {
                    w.append(value.toString());
                }
                w.append("</v>");
            }
            w.append("</c>");
        }

    }
}
