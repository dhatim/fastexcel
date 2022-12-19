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
 * A border element defines the style and color of one or multiple sides of a
 * cell or range of cells.
 */
class BorderElement {

    /**
     * Default border element: no border.
     */
    protected static final BorderElement NONE = new BorderElement(null, null);

    /**
     * Border style.
     */
    private final String style;

    /**
     * RGB border color.
     */
    private final String rgbColor;

    /**
     * Constructor.
     *
     * @param style Border style. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc844549(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @param rgbColor RGB border color.
     */
    BorderElement(String style, String rgbColor) {
        this.style = style;
        this.rgbColor = rgbColor;
    }

    /**
     * Return a new border element with an updated style.
     *
     * @param style Updated style.
     * @return New border element.
     */
    BorderElement updateStyle(String style) {
        return new BorderElement(style, this.rgbColor);
    }

    /**
     * Return a new border element with an updated RGB border color.
     *
     * @param rgbColor Updated RGB border color.
     * @return New border element.
     */
    BorderElement updateColor(String rgbColor) {
        return new BorderElement(this.style, rgbColor);
    }

    @Override
    public int hashCode() {
        return Objects.hash(style, rgbColor);
    }

    @Override
    public boolean equals(Object obj) {
        boolean result;
        if (obj != null && obj.getClass() == this.getClass()) {
            BorderElement other = (BorderElement) obj;
            result = Objects.equals(style, other.style) && Objects.equals(rgbColor, other.rgbColor);
        } else {
            result = false;
        }
        return result;
    }

    /**
     * Write this border element as an XML element.
     *
     * @param name Border element name ("left", "right", etc).
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(String name, Writer w) throws IOException {
        w.append("<").append(name);
        if (style == null && rgbColor == null) {
            w.append("/>");
        } else {
            if (style != null) {
                w.append(" style=\"").append(style).append('\"');
            }
            w.append('>');
            if (rgbColor != null) {
                w.append("<color rgb=\"").append(rgbColor).append("\"/>");
            }
            w.append("</").append(name).append(">");
        }
    }
}
