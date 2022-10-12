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

import java.math.BigDecimal;
import java.util.EnumMap;
import java.util.EnumSet;
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
     * RGB color for shading Nth rows.
     */
    private String shadingFillColor;
    /**
     * Shading row frequency.
     */
    private int eachNRows;
    /**
     * Bold flag.
     */
    private boolean bold;
    /**
     * Italic flag.
     */
    private boolean italic;
    /**
     * Underlined flag.
     */
    private boolean underlined;
    /**
     * Font name.
     */
    private String fontName;
    /**
     * Font size.
     */
    private BigDecimal fontSize;
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
     * Border.
     */
    private Border border;

    /**
     * Protection options.
     */
    private Map<ProtectionOption, Boolean> protectionOptions;

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
     * Shade Nth rows.
     *
     * @param rgb RGB shading color.
     * @param eachNRows shading frequency.
     * @return This style setter.
     */
    public StyleSetter shadeRows(String rgb, int eachNRows) {
        this.shadingFillColor = rgb;
        this.eachNRows = eachNRows;
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
     * Set font name.
     *
     * @param name Font name.
     * @return This style setter.
     */
    public StyleSetter fontName(String name) {
        this.fontName = name;
        return this;
    }

    /**
     * Set font size.
     *
     * @param size Font size, in points.
     * @return This style setter.
     */
    public StyleSetter fontSize(BigDecimal size) {
        this.fontSize = size;
        return this;
    }

    /**
     * Set font size.
     *
     * @param size Font size, in points.
     * @return This style setter.
     */
    public StyleSetter fontSize(int size) {
        this.fontSize = BigDecimal.valueOf(size);
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
     * Use underlined text.
     *
     * @return This style setter.
     */
    public StyleSetter underlined() {
        this.underlined = true;
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
     * Set cell border element.
     *
     * @param side Border side where element is set.
     * @param element Border element to set.
     * @return This style setter.
     */
    private StyleSetter borderElement(BorderSide side, BorderElement element) {
        if (border == null) {
            border = new Border();
        }
        border.setElement(side, element);
        return this;
    }

    /**
     * Apply cell border style on all sides, except diagonal.
     *
     * @param borderStyle Border style.
     * @return This style setter.
     */
    public StyleSetter borderStyle(BorderStyle borderStyle) {
        return borderStyle(borderStyle.xmlValue);
    }

    /**
     * Apply cell border style on all sides, except diagonal.
     *
     * @param borderStyle Border style. Possible values are defined
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borderstylevalues(v=office.14).aspx">here</a>.
     * @return This style setter.
     */
    public StyleSetter borderStyle(String borderStyle) {
        if (border == null) {
            border = new Border();
        }
        EnumSet.of(BorderSide.TOP, BorderSide.LEFT, BorderSide.BOTTOM, BorderSide.RIGHT).forEach(side ->
            borderElement(side, border.elements.get(side).updateStyle(borderStyle))
        );
        return this;
    }

    /**
     * Apply cell border style on a side.
     *
     * @param side Border side where to apply the given style.
     * @param borderStyle Border style.
     * @return This style setter.
     */
    public StyleSetter borderStyle(BorderSide side, BorderStyle borderStyle) {
        return borderStyle(side, borderStyle.xmlValue);
    }
    /**
     * Apply cell border style on a side.
     *
     * @param side Border side where to apply the given style.
     * @param borderStyle Border style. Possible values are defined
     * <a href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borderstylevalues(v=office.14).aspx">here</a>.
     * @return This style setter.
     */
    public StyleSetter borderStyle(BorderSide side, String borderStyle) {
        if (border == null) {
            border = new Border();
        }
        return borderElement(side, border.elements.get(side).updateStyle(borderStyle));
    }

    /**
     * Set cell border color.
     *
     * @param borderColor RGB border color.
     * @return This style setter.
     */
    public StyleSetter borderColor(String borderColor) {
        if (border == null) {
            border = new Border();
        }
        EnumSet.of(BorderSide.TOP, BorderSide.LEFT, BorderSide.BOTTOM, BorderSide.RIGHT).forEach(side ->
            borderElement(side, border.elements.get(side).updateColor(borderColor))
        );
        return this;
    }

    /**
     * Set cell border color.
     *
     * @param side Border side where to apply the given border color.
     * @param borderColor RGB border color.
     * @return This style setter.
     */
    public StyleSetter borderColor(BorderSide side, String borderColor) {
        if (border == null) {
            border = new Border();
        }
        return borderElement(side, border.elements.get(side).updateColor(borderColor));
    }

    /**
     * Sets the value for a protection option.
     *
     * @param option The option to set
     * @param value The value to set for the given option.
     * @return This style setter.
     */
    public StyleSetter protectionOption(ProtectionOption option, Boolean value) {
        if (protectionOptions == null) {
            protectionOptions = new EnumMap<>(ProtectionOption.class);
        }
        protectionOptions.put(option, value);
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
        Alignment alignment;
        if (horizontalAlignment != null || verticalAlignment != null || wrapText) {
            alignment = new Alignment(horizontalAlignment, verticalAlignment, wrapText);
        } else {
            alignment = null;
        }
        Font font;
        if (bold || italic || underlined || fontColor != null || fontName != null || fontSize != null) {
            font = Font.build(bold, italic, underlined, fontName, fontSize, fontColor);
        } else {
            font = Font.DEFAULT;
        }
        Fill fill;
        if (fillColor == null) {
            fill = Fill.NONE;
        } else {
            fill = Fill.fromColor(fillColor);
        }
        if (border == null) {
            border = Border.NONE;
        }

        Protection protection;
        if (protectionOptions != null) {
            protection = new Protection(protectionOptions);
        } else {
            protection = null;
        }

        // Compute a map giving new styles for current styles
        Set<Integer> currentStyles = range.getStyles();
        Map<Integer, Integer> newStyles = currentStyles.stream().collect(Collectors.toMap(Function.identity(), s -> range.getWorksheet().getWorkbook().mergeAndCacheStyle(s, valueFormatting, font, fill, border, alignment, protection)));

        // Apply styles to range
        range.applyStyle(newStyles);

        // Shading color for alternate rows is cached separately
        if (alternateShadingFillColor != null) {
            range.shadeAlternateRows(Fill.fromColor(alternateShadingFillColor, false));
        }

        if (shadingFillColor != null) {
            range.shadeRows(Fill.fromColor(shadingFillColor, false), eachNRows);
        }
    }
    
    /**
     * Apply style elements conditionally
     * @param conditionalFormattingRule Conditional formatting rule to apply
     */
    public void set(ConditionalFormattingRule conditionalFormattingRule) {
        Alignment alignment = null;
        if (horizontalAlignment != null || verticalAlignment != null || wrapText) {
            alignment = new Alignment(horizontalAlignment, verticalAlignment, wrapText);
        }
        Font font = null;
        if (bold || italic || underlined || fontColor != null || fontName != null || fontSize != null) {
            font = Font.build(bold, italic, underlined, fontName, fontSize, fontColor);
        }
        Fill fill = null;
        if (fillColor != null) {
            fill = Fill.fromColor(fillColor, false);
        }
        Protection protection = null;
        if (protectionOptions != null) {
            protection = new Protection(protectionOptions);
        }
        
        int dxfId = range.getWorksheet().getWorkbook().cacheDifferentialFormat(new DifferentialFormat(valueFormatting, font, fill, border, alignment, protection));
        conditionalFormattingRule.setDxfId(dxfId);
        ConditionalFormatting conditionalFormatting = new ConditionalFormatting(range, conditionalFormattingRule);
        range.getWorksheet().addConditionalFormatting(conditionalFormatting);
    }
}
