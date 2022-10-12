package org.dhatim.fastexcel;

import java.io.IOException;
import java.util.Objects;

public class DifferentialFormat {
    private final String valueFormatting;
    private final Font font;
    private final Fill fill;
    private final Border border;
    private final Alignment alignment;
    private final Protection protection;
    private int numFmtId;
    
    /**
     * Constructor.
     *
     * @param numFmtId Id of the value formatting to use
     * @param font Font to use
     * @param fill Fill to use
     * @param border Border to use
     * @param alignment Alignment to use
     * @param protection Proction to use
     */
    DifferentialFormat(String valueFormatting, Font font, Fill fill, Border border, Alignment alignment, Protection protection) {
        this.valueFormatting = valueFormatting;
        this.font = font;
        this.fill = fill;
        this.border = border;
        this.alignment = alignment;
        this.protection = protection;
    }
    
    public String getValueFormatting() {
        return valueFormatting;
    }
    
    public void setNumFmtId(int numFmtId) {
        this.numFmtId = numFmtId;
    }
    
    @Override
    public int hashCode() {
        return Objects.hash(numFmtId, font, fill, border, alignment, protection);
    }

    @Override
    public boolean equals(Object obj) {
        boolean result;
        if (obj != null && obj.getClass() == this.getClass()) {
            DifferentialFormat other = (DifferentialFormat) obj;
            result = Objects.equals(valueFormatting, other.valueFormatting)
                    && Objects.equals(font, other.font)
                    && Objects.equals(fill, other.fill)
                    && Objects.equals(border, other.border)
                    && Objects.equals(alignment, other.alignment)
                    && Objects.equals(protection, other.protection);
        } else {
            result = false;
        }
        return result;
    }

    /**
     * Write this style as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<dxf>");
        if (valueFormatting != null) {
            w.append("<numFmt numFmtId=\"").append(numFmtId).append("\" formatCode=\"").append(valueFormatting).append("\"/>");
        }
        if (font != null) {
            font.write(w);
        }
        if (fill != null) {
            fill.write(w);
        }
        if (border != null) {
            border.write(w);
        }
        if (alignment != null) {
            alignment.write(w);
        }
        if (protection != null) {
            protection.write(w);
        }
        w.append("</dxf>");
    }
}
