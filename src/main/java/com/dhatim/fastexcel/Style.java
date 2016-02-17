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
package com.dhatim.fastexcel;

import java.io.IOException;
import java.util.Objects;

/**
 * A style defines the value formatting, font, fill pattern, border and
 * alignment of a cell or range of cells.
 */
class Style {

    /**
     * Index of cached value formatting.
     */
    private final int valueFormatting;
    /**
     * Index of cached font.
     */
    private final int font;
    /**
     * Index of cached fill pattern.
     */
    private final int fill;
    /**
     * Index of cached border.
     */
    private final int border;
    /**
     * Alignment.
     */
    private final Alignment alignment;

    /**
     * Constructor.
     *
     * @param original Original style. If not {@code null}, its attributes are
     * overridden by the attributes below.
     * @param valueFormatting Index of cached value formatting. Zero if not set.
     * @param font Index of cached font. Zero if not set.
     * @param fill Index of cached fill pattern. Zero if not set.
     * @param border Index of cached border. Zero if not set.
     * @param alignment Alignment. {@code null} if not set.
     */
    Style(Style original, int valueFormatting, int font, int fill, int border, Alignment alignment) {
        this.valueFormatting = (valueFormatting == 0 && original != null) ? original.valueFormatting : valueFormatting;
        this.font = (font == 0 && original != null) ? original.font : font;
        this.fill = (fill == 0 && original != null) ? original.fill : fill;
        this.border = (border == 0 && original != null) ? original.border : border;
        this.alignment = (alignment == null && original != null) ? original.alignment : alignment;
    }

    @Override
    public int hashCode() {
        int hash = 5;
        hash = 29 * hash + this.valueFormatting;
        hash = 29 * hash + this.font;
        hash = 29 * hash + this.fill;
        hash = 29 * hash + this.border;
        hash = 29 * hash + Objects.hashCode(this.alignment);
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final Style other = (Style) obj;
        if (this.valueFormatting != other.valueFormatting) {
            return false;
        }
        if (this.font != other.font) {
            return false;
        }
        if (this.fill != other.fill) {
            return false;
        }
        if (this.border != other.border) {
            return false;
        }
        if (!Objects.equals(this.alignment, other.alignment)) {
            return false;
        }
        return true;
    }

    /**
     * Write this style as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<xf numFmtId=\"").append(valueFormatting).append("\" fontId=\"").append(font).append("\" fillId=\"").append(fill).append("\" borderId=\"").append(border).append("\" xfId=\"0\"");
        if (alignment == null) {
            w.append("/>");
        } else {
            w.append('>');
            alignment.write(w);
            w.append("</xf>");
        }
    }
}
