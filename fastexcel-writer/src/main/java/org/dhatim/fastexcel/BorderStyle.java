package org.dhatim.fastexcel;

public enum BorderStyle {
    NONE("none"),
    THIN("thin"),
    MEDIUM("medium"),
    DASHED("dashed"),
    DOTTED("dotted"),
    THICK("thick"),
    DOUBLE("double"),
    HAIR("hair"),
    MEDIUM_DASHED("mediumDashed"),
    DASH_DOT("dashDot"),
    MEDIUM_DASH_DOT("mediumDashDot"),
    DASH_DOT_DOT("dashDotDot"),
    MEDIUM_DASH_DOT_DOT("mediumDashDotDot"),
    SLANT_DASH_DOT("slantDashDot");

    final String xmlValue;

    BorderStyle(String xmlValue) {
        this.xmlValue = xmlValue;
    }
}
