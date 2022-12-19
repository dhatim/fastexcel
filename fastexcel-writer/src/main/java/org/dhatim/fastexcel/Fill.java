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
 * Pattern fill definition. This does not support any fancy gradient fill.
 */
class Fill {

    /**
     * Reserved Excel fill: none.
     */
    protected static final Fill NONE = new Fill("none", null, true);

    /**
     * Reserved Excel fill: "gray125".
     */
    protected static final Fill GRAY125 = new Fill("gray125", null, true);

    /**
     * Pattern type.
     */
    private final String patternType;
    /**
     * RGB fill color.
     */
    private final String colorRgb;
    /**
     * Foreground/background selection.
     */
    private final boolean fg;

    /**
     * Constructor.
     *
     * @param patternType Pattern type. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc820608(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @param colorRgb RGB pattern color.
     * @param fg Foreground/background selection: {@code true} for solid cell
     * fills and {@code false} for pattern cell fills.
     */
    Fill(String patternType, String colorRgb, boolean fg) {
        this.patternType = patternType;
        this.colorRgb = colorRgb;
        this.fg = fg;
    }

    /**
     * Create a solid pattern fill.
     *
     * @param fgColorRgb RGB fill color.
     * @return New pattern fill.
     */
    static Fill fromColor(String fgColorRgb) {
        return fromColor(fgColorRgb, true);
    }

    /**
     * Create a pattern fill.
     *
     * @param colorRgb RGB fill color.
     * @param fg Foreground/background selection.
     * @return New pattern fill.
     */
    static Fill fromColor(String colorRgb, boolean fg) {
        return new Fill("solid", colorRgb, fg);
    }

    @Override
    public int hashCode() {
        return Objects.hash(patternType, colorRgb, fg);
    }

    @Override
    public boolean equals(Object obj) {
        boolean result;
        if (obj != null && obj.getClass() == this.getClass()) {
            Fill other = (Fill) obj;
            result = Objects.equals(patternType, other.patternType) && Objects.equals(colorRgb, other.colorRgb) && Objects.equals(fg, other.fg);
        } else {
            result = false;
        }
        return result;
    }

    /**
     * Write this fill pattern as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<fill><patternFill patternType=\"").append(patternType).append('\"');
        if (colorRgb == null) {
            w.append("/>");
        } else {
            w.append("><").append(fg ? "fg" : "bg").append("Color rgb=\"").append(colorRgb).append("\"/></patternFill>");
        }
        w.append("</fill>");
    }
}
