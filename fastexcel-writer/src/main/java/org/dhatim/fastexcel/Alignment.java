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
import java.util.Objects;

/**
 * Alignment attributes. For more information refer to
 * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc801104(v=office.14)?redirectedfrom=MSDN">this</a>.
 */
class Alignment {

    private final String horizontal;
    private final String vertical;
    private final boolean wrapText;
    private final int rotation;
    private final int indent;

    /**
     * Constructor.
     *
     * @param horizontal Horizontal alignment attribute. Possible values are
     * defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc880467(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @param vertical Vertical alignment attribute. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc802119(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @param wrapText Enable or disable text wrapping in cells.
     * @param indent Indentation of text in cell
     */
    Alignment(String horizontal, String vertical, boolean wrapText, int rotation, int indent) {
        this.horizontal = horizontal;
        this.vertical = vertical;
        this.wrapText = wrapText;
        this.rotation = rotation;
        this.indent = indent;
    }

    @Override
    public int hashCode() {
        return Objects.hash(horizontal, vertical, wrapText);
    }

    @Override
    public boolean equals(Object obj) {
        boolean result;
        if (obj != null && obj.getClass() == this.getClass()) {
            Alignment other = (Alignment) obj;
            result = Objects.equals(horizontal, other.horizontal) && Objects.equals(vertical, other.vertical)
                    && Objects.equals(wrapText, other.wrapText) && Objects.equals(rotation, other.rotation)
                    && Objects.equals(indent, other.indent);
        } else {
            result = false;
        }
        return result;
    }

    /**
     * Write this alignment as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<alignment");
        if (horizontal != null) {
            w.append(" horizontal=\"").append(horizontal).append('\"');
        }
        if (vertical != null) {
            w.append(" vertical=\"").append(vertical).append('\"');
        }
        if (rotation != 0) {
            w.append(" textRotation=\"").append(rotation).append('\"');
        }
        if (indent != 0) {
            w.append(" indent=\"").append(indent).append('\"');
        }
        if (wrapText) {
            w.append(" wrapText=\"true\"");
        }
        w.append("/>");
    }
}
