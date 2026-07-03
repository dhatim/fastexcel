package org.dhatim.fastexcel;

final class StyleSpec {

    private final String numberingFormat;
    private final Font font;
    private final Fill fill;
    private final Border border;
    private final Alignment alignment;
    private final Protection protection;

    StyleSpec(String numberingFormat, Font font, Fill fill, Border border, Alignment alignment, Protection protection) {
        this.numberingFormat = numberingFormat;
        this.font = font;
        this.fill = fill;
        this.border = border;
        this.alignment = alignment;
        this.protection = protection;
    }

    String getNumberingFormat() {
        return numberingFormat;
    }

    Font getFont() {
        return font;
    }

    Fill getFill() {
        return fill;
    }

    Border getBorder() {
        return border;
    }

    Alignment getAlignment() {
        return alignment;
    }

    Protection getProtection() {
        return protection;
    }
}
