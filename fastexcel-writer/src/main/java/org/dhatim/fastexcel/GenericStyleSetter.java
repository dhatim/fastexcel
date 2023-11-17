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

interface StylesFunction {

    void applyStyles(Map<Integer, Integer> stylesMap);
}

/**
 * Generic helper class to set style elements on a cell or range of cells. This class
 * implements the builder pattern to easily modify a bunch of attributes.<p>
 * For example:
 * <blockquote><pre>
 *  Worksheet ws = ...
 *  ws.range(1, 1, 1, 10).style().borderStyle("thin").bold().fillColor(Color.GRAY4).horizontalAlignment("center").set();
 * </pre></blockquote>
 */
abstract class GenericStyleSetter<STYLE_SETTER extends GenericStyleSetter<STYLE_SETTER>> {

    /**
     * Worksheet in which this styling is applied
     */
    private final Worksheet worksheet;
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
    private Boolean bold;
    /**
     * Italic flag.
     */
    private Boolean italic;
    /**
     * Underlined flag.
     */
    private Boolean underlined;
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
     * Text rotation in degrees
     */
    private int rotation;
    /**
    * Represents the indent level
    */
    private int indent;
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
     * @param worksheet Worksheet in which this styling is applied
     */
    GenericStyleSetter(Worksheet worksheet) {
        this.worksheet = worksheet;
    }

    protected abstract STYLE_SETTER getThis();

    /**
     * Set numbering format.
     *
     * @param numberingFormat Numbering format. For more information, refer to
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee857658(v=office.14)?redirectedfrom=MSDN">this
     * page</a>.
     * @return This style setter.
     */
    public STYLE_SETTER format(String numberingFormat) {
        this.valueFormatting = numberingFormat;
        return getThis();
    }

    /**
     * Set fill color.
     *
     * @param rgb RGB fill color. See {@link Color} for predefined values.
     * @return This style setter.
     */
    public STYLE_SETTER fillColor(String rgb) {
        this.fillColor = rgb;
        return getThis();
    }

    /**
     * Shade alternate rows.
     *
     * @param rgb RGB shading color.
     * @return This style setter.
     */
    public STYLE_SETTER shadeAlternateRows(String rgb) {
        this.alternateShadingFillColor = rgb;
        return getThis();
    }

    /**
     * Shade Nth rows.
     *
     * @param rgb RGB shading color.
     * @param eachNRows shading frequency.
     * @return This style setter.
     */
    public STYLE_SETTER shadeRows(String rgb, int eachNRows) {
        this.shadingFillColor = rgb;
        this.eachNRows = eachNRows;
        return getThis();
    }

    /**
     * Set font color.
     *
     * @param rgb RGB font color.
     * @return This style setter.
     */
    public STYLE_SETTER fontColor(String rgb) {
        this.fontColor = rgb;
        return getThis();
    }

    /**
     * Set font name.
     *
     * @param name Font name.
     * @return This style setter.
     */
    public STYLE_SETTER fontName(String name) {
        this.fontName = name;
        return getThis();
    }

    /**
     * Set font size.
     *
     * @param size Font size, in points.
     * @return This style setter.
     */
    public STYLE_SETTER fontSize(BigDecimal size) {
        this.fontSize = size;
        return getThis();
    }

    /**
     * Set font size.
     *
     * @param size Font size, in points.
     * @return This style setter.
     */
    public STYLE_SETTER fontSize(int size) {
        this.fontSize = BigDecimal.valueOf(size);
        return getThis();
    }

    /**
     * Use bold text.
     *
     * @return This style setter.
     */
    public STYLE_SETTER bold() {
        this.bold = true;
        return getThis();
    }

    /**
     * Use italic text.
     *
     * @return This style setter.
     */
    public STYLE_SETTER italic() {
        this.italic = true;
        return getThis();
    }

    /**
     * Use underlined text.
     *
     * @return This style setter.
     */
    public STYLE_SETTER underlined() {
        this.underlined = true;
        return getThis();
    }

    /**
     * Define horizontal alignment.
     *
     * @param alignment Horizontal alignment. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc880467(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @return This style setter.
     */
    public STYLE_SETTER horizontalAlignment(String alignment) {
        this.horizontalAlignment = alignment;
        return getThis();
    }

    /**
     * Define vertical alignment.
     *
     * @param alignment Vertical alignment. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc802119(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @return This style setter.
     */
    public STYLE_SETTER verticalAlignment(String alignment) {
        this.verticalAlignment = alignment;
        return getThis();
    }

    /**
     * Enable or disable text wrapping in cells.
     *
     * @param wrapText {@code true} to enable text wrapping (default is
     * {@code false}).
     * @return This style setter.
     */
    public STYLE_SETTER wrapText(boolean wrapText) {
        this.wrapText = wrapText;
        return getThis();
    }

    /**
     * Set cell text rotation in degrees.
     *
     * @param degrees rotation of text in cell
     * @return This style setter
     */
    public STYLE_SETTER rotation(int degrees) {
        this.rotation = degrees;
        return getThis();
    }

     /**
     * Set cell text indentation.
     *
     * @param indent indentation of text in cell
     * @return This style setter
     */
    public STYLE_SETTER indent(int indent) {
        if (indent < 0) {
            this.indent = 0;
        } else if (indent > 15) {
            this.indent = 15;
        }else{
            this.indent = indent;
        }
        return getThis();
    }

    /**
     * Set cell border element.
     *
     * @param side Border side where element is set.
     * @param element Border element to set.
     * @return This style setter.
     */
    private STYLE_SETTER borderElement(BorderSide side, BorderElement element) {
        if (border == null) {
            border = new Border();
        }
        border.setElement(side, element);
        return getThis();
    }

    /**
     * Apply cell border style on all sides, except diagonal.
     *
     * @param borderStyle Border style.
     * @return This style setter.
     */
    public STYLE_SETTER borderStyle(BorderStyle borderStyle) {
        return borderStyle(borderStyle.xmlValue);
    }

    /**
     * Apply cell border style on all sides, except diagonal.
     *
     * @param borderStyle Border style. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.borderstylevalues?view=openxml-2.8.1">here</a>.
     * @return This style setter.
     */
    public STYLE_SETTER borderStyle(String borderStyle) {
        if (border == null) {
            border = new Border();
        }
        EnumSet.of(BorderSide.TOP, BorderSide.LEFT, BorderSide.BOTTOM, BorderSide.RIGHT).forEach(side ->
                borderElement(side, border.elements.get(side).updateStyle(borderStyle))
        );
        return getThis();
    }

    /**
     * Apply cell border style on a side.
     *
     * @param side Border side where to apply the given style.
     * @param borderStyle Border style.
     * @return This style setter.
     */
    public STYLE_SETTER borderStyle(BorderSide side, BorderStyle borderStyle) {
        return borderStyle(side, borderStyle.xmlValue);
    }
    /**
     * Apply cell border style on a side.
     *
     * @param side Border side where to apply the given style.
     * @param borderStyle Border style. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc844549(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @return This style setter.
     */
    public STYLE_SETTER borderStyle(BorderSide side, String borderStyle) {
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
    public STYLE_SETTER borderColor(String borderColor) {
        if (border == null) {
            border = new Border();
        }
        EnumSet.of(BorderSide.TOP, BorderSide.LEFT, BorderSide.BOTTOM, BorderSide.RIGHT).forEach(side ->
                borderElement(side, border.elements.get(side).updateColor(borderColor))
        );
        return getThis();
    }

    /**
     * Set cell border color.
     *
     * @param side Border side where to apply the given border color.
     * @param borderColor RGB border color.
     * @return This style setter.
     */
    public STYLE_SETTER borderColor(BorderSide side, String borderColor) {
        if (border == null) {
            border = new Border();
        }
        return borderElement(side, border.elements.get(side).updateColor(borderColor));
    }

    /**
     * Set cell diagonal property.
     *
     * @param diagonalProperty Diagonal border property which should be aplied to a cell
     * @return This style setter.
     */
    public STYLE_SETTER diagonalProperty(DiagonalProperty diagonalProperty) {
        if (border == null) {
            border = new Border();
        }
        border.setDiagonalProperty(diagonalProperty);
        return getThis();
    }

    /**
     * Sets the value for a protection option.
     *
     * @param option The option to set
     * @param value The value to set for the given option.
     * @return This style setter.
     */
    public STYLE_SETTER protectionOption(ProtectionOption option, Boolean value) {
        if (protectionOptions == null) {
            protectionOptions = new EnumMap<>(ProtectionOption.class);
        }
        protectionOptions.put(option, value);
        return getThis();
    }

    /**
     * Apply style elements. <b>Do not forget to call this method when you are
     * done otherwise style changes are lost!</b>
     */
    public abstract void set();

    /**
     * Apply style elements. <b>Do not forget to call this method when you are
     * done otherwise style changes are lost!</b>
     *
     * @param shadingEnabled Whether shading should be enabled in this styling
     * @param currentStyles Set of current styles from styled property
     * @param stylesFunction A function which applies merged newStyles to styled property
     */
    protected void setStyle(boolean shadingEnabled, Set<Integer> currentStyles,
                    StylesFunction stylesFunction) {
        Alignment alignment;
        if (horizontalAlignment != null || verticalAlignment != null || wrapText || rotation != 0 || indent  != 0) {
            alignment = new Alignment(horizontalAlignment, verticalAlignment, wrapText, rotation, indent);
        } else {
            alignment = null;
        }
        Font font;
        if (!Font.equalsDefault(bold,italic,underlined,fontName,fontSize,fontColor)) {
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
        Map<Integer, Integer> newStyles = currentStyles.stream().collect(Collectors.toMap(Function.identity(), s -> worksheet.getWorkbook().mergeAndCacheStyle(s, valueFormatting, font, fill, border, alignment, protection)));

        // Apply styles
        stylesFunction.applyStyles(newStyles);

        if (shadingEnabled) {
            // Shading color for alternate rows is cached separately
            if (alternateShadingFillColor != null) {
                getRange().shadeAlternateRows(Fill.fromColor(alternateShadingFillColor, false));
            }

            if (shadingFillColor != null) {
                getRange().shadeRows(Fill.fromColor(shadingFillColor, false), eachNRows);
            }
        }
    }

    /**
     * Apply style elements conditionally
     * @param conditionalFormattingRule Conditional formatting rule to apply
     */
    public void set(ConditionalFormattingRule conditionalFormattingRule) {
        Alignment alignment = null;
        if (horizontalAlignment != null || verticalAlignment != null || wrapText || rotation != 0 || indent != 0) {
            alignment = new Alignment(horizontalAlignment, verticalAlignment, wrapText, rotation, indent);
        }
        Font font = null;
        if (bold != null && bold || italic != null && italic || underlined != null && underlined || fontColor != null || fontName != null || fontSize != null) {
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

        int dxfId = worksheet.getWorkbook().cacheDifferentialFormat(new DifferentialFormat(valueFormatting, font, fill, border, alignment, protection));
        conditionalFormattingRule.setDxfId(dxfId);
        ConditionalFormatting conditionalFormatting = new ConditionalFormatting(getRange(), conditionalFormattingRule);
        worksheet.addConditionalFormatting(conditionalFormatting);
    }

    protected abstract Range getRange();
}
