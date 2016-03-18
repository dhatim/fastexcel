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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;

/**
 * A cell contains a value and a cached style index.
 */
class Cell {

    /**
     * Cell value.
     */
    private Object value;

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
            w.append("<c r=\"").append(Range.colToString(c)).append(r + 1).append('\"');
            if (style != 0) {
                w.append(" s=\"").append(style).append('\"');
            }
            if (value != null && !(value instanceof Formula)) {
                w.append(" t=\"").append((value instanceof CachedString) ? 's' : 'n').append('\"');
            }
            w.append(">");
            if (value instanceof Formula) {
                w.append("<f>").append(((Formula) value).getExpression()).append("</f>");
            } else if (value != null) {
                w.append("<v>");
                if (value instanceof CachedString) {
                    w.append(((CachedString) value).getIndex());
                } else if (value instanceof Integer) {
                    w.append((int) value);
                } else if (value instanceof Long) {
                    w.append((long) value);
                } else if (value instanceof Double) {
                    w.append((double) value);
                } else {
                    w.append(value.toString());
                }
                w.append("</v>");
            }
            w.append("</c>");
        }
    }

    /**
     * Set value of this cell.
     *
     * @param wb Parent workbook.
     * @param v Cell value. Supported types are
     * {@link String}, {@link Date}, {@link LocalDate}, {@link LocalDateTime}, {@link ZonedDateTime}
     * and {@link Number} implementations. Note Excel timestamps do not carry
     * any timezone information; {@link Date} values are converted to an Excel
     * serial number with the system timezone. If you need a specific timezone,
     * prefer passing a {@link ZonedDateTime}.
     */
    void setValue(Workbook wb, Object v) {
        if (v instanceof String) {
            value = wb.cacheString((String) v);
        } else if (v == null || v instanceof Number) {
            value = v;
        } else if (v instanceof Date) {
            value = TimestampUtil.convertDate((Date) v);
        } else if (v instanceof LocalDateTime) {
            value = TimestampUtil.convertDate(Date.from(((LocalDateTime) v).atZone(ZoneId.systemDefault()).toInstant()));
        } else if (v instanceof LocalDate) {
            value = TimestampUtil.convertDate((LocalDate) v);
        } else if (v instanceof ZonedDateTime) {
            value = TimestampUtil.convertZonedDateTime((ZonedDateTime) v);
        } else {
            throw new IllegalArgumentException("No supported cell type for " + v.getClass());
        }
    }

    /**
     * Get value or formula stored in this cell.
     *
     * @return Value or {@link Formula}, or {@code null}.
     */
    Object getValue() {
        Object result;
        if (value instanceof CachedString) {
            result = ((CachedString) value).getString();
        } else {
            result = value;
        }
        return result;
    }

    /**
     * Assign a formula to this cell.
     *
     * @param expression Formula expression.
     */
    void setFormula(String expression) {
        value = new Formula(expression);
    }

    /**
     * Get the style of this cell.
     *
     * @return Cell style.
     */
    public int getStyle() {
        return style;
    }

    /**
     * Set the style of this cell.
     *
     * @param style New cell style.
     */
    public void setStyle(int style) {
        this.style = style;
    }

}
