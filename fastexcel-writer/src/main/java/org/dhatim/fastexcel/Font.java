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
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Objects;

/**
 * Font definition. For details, check
 * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc862295(v=office.14)?redirectedfrom=MSDN">this
 * page</a>.
 */
class Font {

    /**
     * Default font.
     */
    public static Font DEFAULT = build(false, false, false, "Calibri", BigDecimal.valueOf(11.0), "FF000000");

    /**
     * Bold flag.
     */
    private final boolean bold;
    /**
     * Italic flag.
     */
    private final boolean italic;
    /**
     * Underlined flag.
     */
    private final boolean underlined;
    /**
     * Font name.
     */
    private final String name;
    /**
     * Font size.
     */
    private final BigDecimal size;
    /**
     * RGB font color.
     */
    private final String rgbColor;

    /**
     * Constructor.
     *
     * @param bold Bold flag.
     * @param italic Italic flag.
     * @param underlined Underlined flag.
     * @param name Font name.
     * @param size Font size, in points.
     * @param rgbColor RGB font color.
     */
    Font(boolean bold, boolean italic, boolean underlined, String name, BigDecimal size, String rgbColor) {
        if (size.compareTo(BigDecimal.valueOf(409)) > 0 || size.compareTo(BigDecimal.valueOf(1)) < 0) {
            throw new IllegalStateException("Font size must be between 1 and 409 points: " + size);
        }
        this.bold = bold;
        this.italic = italic;
		this.underlined = underlined;
        this.name = name;
        this.size = size.setScale(2, RoundingMode.HALF_UP);
        this.rgbColor = rgbColor;
    }

    /**
     * Helper to create a new font.
     *
     * @param bold Bold flag.
     * @param italic Italic flag.
     * @param underlined Underlined flag.
     * @param name Font name. Defaults to "Calibri".
     * @param size Font size, in points. Defaults to 11.0.
     * @param rgbColor RGB font color. Defaults to "FF000000".
     * @return New font object.
     */
    public static Font build(Boolean bold, Boolean italic, Boolean underlined, String name, BigDecimal size, String rgbColor) {
        return new Font(bold != null? bold : DEFAULT.bold, italic != null ? italic : DEFAULT.italic , underlined != null ? underlined : DEFAULT.underlined, name != null ? name : DEFAULT.name, size != null ?  size:DEFAULT.size, rgbColor != null ?  rgbColor: DEFAULT.rgbColor);
    }

    @Override
    public int hashCode() {
        return Objects.hash(bold, italic, underlined, name, size, rgbColor);
    }

    @Override
    public boolean equals(Object obj) {
        boolean result;
        if (obj != null && obj.getClass() == this.getClass()) {
            Font other = (Font) obj;
            result = Objects.equals(bold, other.bold) && Objects.equals(italic, other.italic) && Objects.equals(underlined, other.underlined) && Objects.equals(name, other.name) && Objects.equals(size, other.size) && Objects.equals(rgbColor, other.rgbColor);
        } else {
            result = false;
        }
        return result;
    }

    public static boolean equalsDefault(Boolean bold, Boolean italic, Boolean underlined, String fontName, BigDecimal fontSize, String fontColor) {
        return Objects.equals(bold, DEFAULT.bold) && Objects.equals(italic, DEFAULT.italic) && Objects.equals(underlined, DEFAULT.underlined) && Objects.equals(fontName, DEFAULT.name) && Objects.equals(fontSize, DEFAULT.size) && Objects.equals(fontColor, DEFAULT.rgbColor);
    }

    /**
     * Write this font as an XML element.
     *
     * @param w Output writer
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<font>").append(bold ? "<b/>" : "").append(italic ? "<i/>" : "").append(underlined ? "<u/>" : "").append("<sz val=\"").append(size.toString()).append("\"/>");
        if (rgbColor != null) {
            w.append("<color rgb=\"").append(rgbColor).append("\"/>");
        }
        w.append("<name val=\"").appendEscaped(name).append("\"/>");
        w.append("</font>");
    }
}
