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
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.util.Date;

import static org.dhatim.fastexcel.CellAddress.convertNumToColString;

/**
 * A cell contains a value and a cached style index.
 */
class Cell implements Ref {

    /**
     * Cell value.
     */
    private CellValue value;

    /**
     * Cached style index.
     */
    private int style;

    /**
     * Write this cell as an XML element.
     *
     * @param w Output writer.
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w, int r, int c) throws IOException {
        if (value != null || style != 0) {
            w.append("<c r=\"").append(convertNumToColString(c)).append(r + 1).append('\"');
            if (style != 0) {
                w.append(" s=\"").append(style).append('\"');
            }
            if (value != null && value.type() != null) {
                w.append(" t=\"").append(value.type()).append('\"');
            }
            w.append(">");
            if (value != null) {
                value.write(w);
            }
            w.append("</c>");
        }
    }

    void setValue(Workbook wb, String v) {
        value = v == null ? null : new CachedStringValue(wb.cacheString(v));
    }

    void setValue(Number v) {
        value = v == null ? null : new NumericValue(v);
    }

    void setValue(Boolean v) {
        value = v == null ? null : new BooleanValue(v);
    }

    void setValue(Date v) {
        setValue(v == null ? null : TimestampUtil.convertDate(v));
    }

    void setValue(LocalDateTime v) {
        setValue(v == null ? null : TimestampUtil.convertDate(v));
    }

    void setValue(LocalDate v) {
        setValue(v == null ? null : TimestampUtil.convertDate(v));
    }

    void setValue(ZonedDateTime v) {
        setValue(v == null ? null : TimestampUtil.convertZonedDateTime(v));
    }

    void setValue(Instant v) {
        setValue(v == null ? null : TimestampUtil.convertInstant(v));
    }

    /**
     * Get value or formula stored in this cell.
     *
     * @return Value or {@link Formula}, or {@code null}.
     */
    Object getValue() {
        return value == null ? null : value.value();
    }

    /**
     * Assign a formula to this cell.
     *
     * @param expression Formula expression.
     */
    void setFormula(String expression) {
        value = new FormulaValue(new Formula(expression));
    }

    /**
     * Assign an inline string to this cell.
     *
     * @param v String value.
     */
    void setInlineString(String v) {
        value = v == null ? null : new InlineStringValue(v);
    }

    /**
     * Assign a rich inline string to this cell.
     *
     * @param v Rich inline string value.
     */
    void setInlineString(RichText v) {
        value = v == null ? null : new RichTextValue(v);
    }

    /**
     * Get the style of this cell.
     *
     * @return Cell style.
     */
    int getStyle() {
        return style;
    }

    /**
     * Set the style of this cell.
     *
     * @param style New cell style.
     */
    void setStyle(int style) {
        this.style = style;
    }

    private interface CellValue {
        String type();

        void write(Writer w) throws IOException;

        Object value();
    }

    private static final class FormulaValue implements CellValue {
        private final Formula formula;

        private FormulaValue(Formula formula) {
            this.formula = formula;
        }

        public String type() {
            return null;
        }

        public void write(Writer w) throws IOException {
            w.append("<f>").append(formula.getExpression()).append("</f>");
        }

        public Object value() {
            return formula;
        }
    }

    private static final class RichTextValue implements CellValue {
        private final RichText richText;

        private RichTextValue(RichText richText) {
            this.richText = richText;
        }

        public String type() {
            return "inlineStr";
        }

        public void write(Writer w) throws IOException {
            richText.write(w);
        }

        public Object value() {
            return richText;
        }
    }

    private static final class InlineStringValue implements CellValue {
        private final String string;

        private InlineStringValue(String string) {
            this.string = string;
        }

        public String type() {
            return "inlineStr";
        }

        public void write(Writer w) throws IOException {
            w.append("<is><t>").appendEscaped(string).append("</t></is>");
        }

        public Object value() {
            return string;
        }
    }

    private static final class CachedStringValue implements CellValue {
        private final CachedString cachedString;

        private CachedStringValue(CachedString cachedString) {
            this.cachedString = cachedString;
        }

        public String type() {
            return "s";
        }

        public void write(Writer w) throws IOException {
            w.append("<v>").append(cachedString.getIndex()).append("</v>");
        }

        public Object value() {
            return cachedString.getString();
        }
    }

    private static final class NumericValue implements CellValue {
        private final Number number;

        private NumericValue(Number number) {
            this.number = number;
        }

        public String type() {
            return "n";
        }

        public void write(Writer w) throws IOException {
            w.append("<v>").append(number.toString()).append("</v>");
        }

        public Object value() {
            return number;
        }
    }

    private static final class BooleanValue implements CellValue {
        private final Boolean bool;

        private BooleanValue(Boolean bool) {
            this.bool = bool;
        }

        public String type() {
            return "b";
        }

        public void write(Writer w) throws IOException {
            w.append("<v>").append(bool ? '1' : '0').append("</v>");
        }

        public Object value() {
            return bool;
        }
    }

}
