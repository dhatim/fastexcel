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

import java.util.Collections;
import java.util.HashSet;

import static org.dhatim.fastexcel.Worksheet.MAX_ROWS;

/**
 * Helper class to set style elements on a column. This class
 * implements the builder pattern to easily modify a bunch of attributes.<p>
 * For example:
 * <blockquote><pre>
 *  Worksheet ws = ...
 *  ws.range(1, 1, 1, 10).style().borderStyle("thin").bold().fillColor(Color.GRAY4).horizontalAlignment("center").set();
 * </pre></blockquote>
 */
public class ColumnStyleSetter extends GenericStyleSetter<ColumnStyleSetter> {

    /**
     * Column where the style is applied.
     */
    private final Column column;

    /**
     * Constructor.
     *
     * @param column Column where style is modified.
     */
    ColumnStyleSetter(Column column) {
        super(column.getWorksheet());
        this.column = column;
    }

    /**
     * Apply style elements. <b>Do not forget to call this method when you are
     * done otherwise style changes are lost!</b>
     */
    public void set() {
        super.setStyle(false, new HashSet<>(Collections.singletonList(column.getStyle())), column::applyStyle);
    }

    @Override
    protected Range getRange() {
        int colNumber = column.getColNumber();
        return column.getWorksheet().range(0, colNumber, MAX_ROWS - 1, colNumber);
    }

    @Override
    protected ColumnStyleSetter getThis() {
        return this;
    }
}
