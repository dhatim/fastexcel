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

import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Helper class to set style elements on a cell or range of cells. This class
 * implements the builder pattern to easily modify a bunch of attributes.<p>
 * For example:
 * <blockquote><pre>
 *  Worksheet ws = ...
 *  ws.range(1, 1, 1, 10).style().borderStyle("thin").bold().fillColor(Color.GRAY4).horizontalAlignment("center").set();
 * </pre></blockquote>
 */
public class StyleSetter {

    /**
     * Range of cells where the style is applied.
     */
    private final Range range;
    /**
     * Value formatting.
     */
    private String valueFormatting;
    /**
     * RGB fill color.
     */
    private String fillColor;
    /**
     * RGB color for shading of alternate rows.
     */
    private String alternateShadingFillColor;
    /**
     * Bold flag.
     */
    private boolean bold;
    /**
     * Italic flag.
     */
    private boolean italic;
    /**
     * RGB font color.
     */
    private String fontColor;
    /**
     * Horizontal alignment.
     */
    private String horizontalAlignment;
    /**
     * Vertical alignment.
     */
    private String verticalAlignment;
    /**
     * Wrap text flag.
     */
    private boolean wrapText;
    /**
     * Border style.
     */
    private String borderStyle;
    /**
     * RGB border color.
     */
    private String borderColor;

    /**
     * Constructor.
     *
     * @param range Range of cells where style is modified.
     */
    StyleSetter(Range range) {
        this.range = range;
    }

    /**
     * Set numbering format.
     *
     * @param numberingFormat Numbering format. For more information, refer to
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx">this
     * page</a>.
     * @return This style setter.
     */
    public StyleSetter format(String numberingFormat) {
        this.valueFormatting = numberingFormat;
        return this;
    }

    /**
     * Set fill color.
     *
     * @param rgb RGB fill color. See {@link Color} for predefined values.
     * @return This style setter.
     */
    public StyleSetter fillColor(String rgb) {
        this.fillColor = rgb;
        return this;
    }

    /**
     * Shade alternate rows.
     *
     * @param rgb RGB shading color.
     * @return This style setter.
     */
    public StyleSetter shadeAlternateRows(String rgb) {
        this.alternateShadingFillColor = rgb;
        return this;
    }

    /**
     * Set font color.
     *
     * @param rgb RGB font color.
     * @return This style setter.
     */
    public StyleSetter fontColor(String rgb) {
        this.fontColor = rgb;
        return this;
    }

    /**
     * Use bold text.
     *
     * @return This style setter.
     */
    public StyleSetter bold() {
        this.bold = true;
        return this;
    }

    /**
     * Use italic text.
     *
     * @return This style setter.
     */
    public StyleSetter italic() {
        this.italic = true;
        return this;
    }

    /**
     * Define horizontal alignment.
     *
     * @param alignment Horizontal alignment. Possible values are defined
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.horizontalalignmentvalues(v=office.14).aspx">here</a>.
     * @return This style setter.
     */
    public StyleSetter horizontalAlignment(String alignment) {
        this.horizontalAlignment = alignment;
        return this;
    }

    /**
     * Define vertical alignment.
     *
     * @param alignment Vertical alignment. Possible values are defined
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.verticalalignmentvalues(v=office.14).aspx">here</a>.
     * @return This style setter.
     */
    public StyleSetter verticalAlignment(String alignment) {
        this.verticalAlignment = alignment;
        return this;
    }

    /**
     * Enable or disable text wrapping in cells.
     *
     * @param wrapText {@code true} to enable text wrapping (default is
     * {@code false}).
     * @return This style setter.
     */
    public StyleSetter wrapText(boolean wrapText) {
        this.wrapText = wrapText;
        return this;
    }

    /**
     * Set cell border style.
     *
     * @param borderStyle Border style. Possible values are defined
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borderstylevalues(v=office.14).aspx">here</a>.
     * @return This style setter.
     */
    public StyleSetter borderStyle(String borderStyle) {
        this.borderStyle = borderStyle;
        return this;
    }

    /**
     * Set cell border color.
     *
     * @param borderColor RGB border color.
     * @return This style setter.
     */
    public StyleSetter borderColor(String borderColor) {
        this.borderColor = borderColor;
        return this;
    }

    /**
     * Merge cells in this style setter's range.
     *
     * @return This style setter.
     */
    public StyleSetter merge() {
        range.merge();
        return this;
    }

    /**
     * Apply style elements. <b>Do not forget to call this method when you are
     * done otherwise style changes are lost!</b>
     */
    public void set() {
        final Alignment alignment;
        if (horizontalAlignment != null || verticalAlignment != null || wrapText == true) {
            alignment = new Alignment(horizontalAlignment, verticalAlignment, wrapText);
        } else {
            alignment = null;
        }
        final Font font;
        if (bold || italic || fontColor != null) {
            font = Font.build(bold, italic, fontColor);
        } else {
            font = Font.DEFAULT;
        }
        final Fill fill;
        if (fillColor != null) {
            fill = Fill.fromColor(fillColor);
        } else {
            fill = Fill.NONE;
        }
        final Border border;
        if (borderStyle != null || borderColor != null) {
            border = Border.fromStyleAndColor(borderStyle, borderColor);
        } else {
            border = Border.NONE;
        }

        // Compute a map giving new styles for current styles
        Set<Integer> currentStyles = range.getStyles();
        Map<Integer, Integer> newStyles = currentStyles.stream().collect(Collectors.toMap(Function.identity(), s -> range.getWorksheet().getWorkbook().mergeAndCacheStyle(s, valueFormatting, font, fill, border, alignment)));

        // Apply styles to range
        range.applyStyle(newStyles);

        // Shading color for alternate rows is cached separately
        if (alternateShadingFillColor != null) {
            range.shadeAlternateRows(Fill.fromColor(alternateShadingFillColor, false));
        }
    }
}
