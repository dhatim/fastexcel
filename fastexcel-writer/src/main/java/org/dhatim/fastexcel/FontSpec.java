package org.dhatim.fastexcel;

import java.math.BigDecimal;

final class FontSpec {

    private final Boolean bold;
    private final Boolean italic;
    private final Boolean underlined;
    private final String name;
    private final BigDecimal size;
    private final String rgbColor;
    private final Boolean strikethrough;

    FontSpec(Boolean bold, Boolean italic, Boolean underlined, String name, BigDecimal size, String rgbColor, Boolean strikethrough) {
        this.bold = bold;
        this.italic = italic;
        this.underlined = underlined;
        this.name = name;
        this.size = size;
        this.rgbColor = rgbColor;
        this.strikethrough = strikethrough;
    }

    Font toFont(Font defaults) {
        return Font.build(defaults, bold, italic, underlined, name, size, rgbColor, strikethrough);
    }

    boolean hasDifferentialOverrides() {
        return Boolean.TRUE.equals(bold) || Boolean.TRUE.equals(italic) || Boolean.TRUE.equals(underlined)
                || Boolean.TRUE.equals(strikethrough) || name != null || size != null || rgbColor != null;
    }
}
