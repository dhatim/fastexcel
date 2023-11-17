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

/**
 * Helper class to set style elements on a cell or range of cells. This class
 * implements the builder pattern to easily modify a bunch of attributes.<p>
 * For example:
 * <blockquote><pre>
 *  Worksheet ws = ...
 *  ws.range(1, 1, 1, 10).style().borderStyle("thin").bold().fillColor(Color.GRAY4).horizontalAlignment("center").set();
 * </pre></blockquote>
 */
public class StyleSetter extends GenericStyleSetter<StyleSetter>{

    /**
     * Range of cells where the style is applied.
     */
    private final Range range;

    /**
     * Constructor.
     *
     * @param range Range of cells where style is modified.
     */
    StyleSetter(Range range) {
        super(range.getWorksheet());
        this.range = range;
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
        super.setStyle(true, range.getStyles(), range::applyStyle);
    }

    @Override
    protected Range getRange() {
        return this.range;
    }

    @Override
    protected StyleSetter getThis() {
        return this;
    }
}
