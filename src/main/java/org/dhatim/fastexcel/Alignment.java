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
 * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.alignment(v=office.14).aspx">this</a>.
 */
class Alignment {

    private final String horizontal;
    private final String vertical;
    private final boolean wrapText;

    /**
     * Constructor.
     *
     * @param horizontal Horizontal alignment attribute. Possible values are
     * defined
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.horizontalalignmentvalues(v=office.14).aspx">here</a>.
     * @param vertical Vertical alignment attribute. Possible values are defined
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.verticalalignmentvalues(v=office.14).aspx">here</a>.
     * @param wrapText Enable or disable text wrapping in cells.
     */
    Alignment(String horizontal, String vertical, boolean wrapText) {
        this.horizontal = horizontal;
        this.vertical = vertical;
        this.wrapText = wrapText;
    }

    @Override
    public int hashCode() {
        return Objects.hash(horizontal, vertical, wrapText);
    }

    @Override
    public boolean equals(Object obj) {
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final Alignment other = (Alignment) obj;
        if (!Objects.equals(this.horizontal, other.horizontal)) {
            return false;
        }
        if (!Objects.equals(this.vertical, other.vertical)) {
            return false;
        }
        if (this.wrapText != other.wrapText) {
            return false;
        }
        return true;
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
        if (wrapText) {
            w.append(" wrapText=\"").append(Boolean.toString(wrapText)).append('\"');
        }
        w.append("/>");
    }
}
