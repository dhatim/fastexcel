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
import java.time.ZonedDateTime;
import java.util.Date;

/**
 * A cell contains a value and a cached style index.
 */
class Cell implements Ref {

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
            w.append("<c r=\"").append(colToString(c)).append(r + 1).append('\"');
            if (style != 0) {
                w.append(" s=\"").append(style).append('\"');
            }
            if (value != null && !(value instanceof Formula)) {
                w.append(" t=\"").append(getCellType(value)).append('\"');
            }
            w.append(">");
            if (value instanceof Formula) {
                w.append("<f>").append(((Formula) value).getExpression()).append("</f>");
            } else if (value instanceof String) {
                w.append("<is><t>").appendEscaped((String) value).append("</t></is>");
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
                } else if (value instanceof Boolean) {
                    w.append((Boolean) value ? '1' : '0');
                } else {
                    w.append(value.toString());
                }
                w.append("</v>");
            }
            w.append("</c>");
        }
    }

    static String getCellType(Object value) {
        if (value instanceof CachedString) {
            return "s";
        } else if (value instanceof Boolean) {
            return "b";
        } else if (value instanceof String) {
            return "inlineStr";
        } else {
            return "n";
        }
    }

    void setValue(Workbook wb, String v) {
        value = v == null ? null : wb.cacheString(v);
    }

    void setValue(Number v) {
        value = v;
    }

    void setValue(Boolean v) {
        value = v;
    }
    void setValue(Date v) {
        value = v == null ? null : TimestampUtil.convertDate(v);
    }

    void setValue(LocalDateTime v) {
        value = v == null ? null :
            TimestampUtil.convertDate(v);
    }

    void setValue(LocalDate v) {
        value = v == null ? null : TimestampUtil.convertDate(v);

    }

    void setValue(ZonedDateTime v) {
        value = v == null ? null : TimestampUtil.convertZonedDateTime(v);
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
     * Assign an inline string to this cell.
     *
     * @param v String value.
     */
    void setInlineString(String v) {
        value = v;
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

}
