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
import java.util.*;

/**
 * Border attributes.
 */
class Border {

    /**
     * Default border attributes: no borders.
     */
    protected static final Border NONE = new Border();

    /**
     * Border elements.
     */
    final Map<BorderSide, BorderElement> elements = new EnumMap<>(BorderSide.class);
    /**
     * Diagonal properties
     */
    private final Set<DiagonalProperty> diagonalProperties = new HashSet<>();

    /**
     * Default constructor.
     */
    Border() {
        this(BorderElement.NONE, BorderElement.NONE, BorderElement.NONE, BorderElement.NONE, BorderElement.NONE);
    }

    /**
     * Simple constructor.
     *
     * @param element Border element for all sides, except diagonal.
     */
    Border(BorderElement element) {
        this(element, element, element, element, BorderElement.NONE);
    }

    /**
     * Constructor.
     *
     * @param left Border element for left side.
     * @param right Border element for right side.
     * @param top Border element for top side.
     * @param bottom Border element for bottom side.
     * @param diagonal Border element for diagonal side.
     */
    Border(BorderElement left, BorderElement right, BorderElement top, BorderElement bottom, BorderElement diagonal) {
        elements.put(BorderSide.TOP, top);
        elements.put(BorderSide.LEFT, left);
        elements.put(BorderSide.BOTTOM, bottom);
        elements.put(BorderSide.RIGHT, right);
        elements.put(BorderSide.DIAGONAL, diagonal);
    }

    /**
     * Set a single border element.
     *
     * @param side Border side.
     * @param element Border element.
     */
    void setElement(BorderSide side, BorderElement element) {
        elements.put(side, element);
    }

    /**
     * Set a diagonal property.
     *
     * @param diagonalProperty Diagonal property.
     */
    void setDiagonalProperty(DiagonalProperty diagonalProperty) {
        diagonalProperties.add(diagonalProperty);
    }

    /**
     * Create a border where all sides have the same style and color.
     *
     * @param style Border style. Possible values are defined
     * <a href="https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc844549(v=office.14)?redirectedfrom=MSDN">here</a>.
     * @param color RGB border color.
     * @return A new border object.
     */
    static Border fromStyleAndColor(String style, String color) {
        BorderElement element = new BorderElement(style, color);
        return new Border(element, element, element, element, BorderElement.NONE);
    }

    @Override
    public int hashCode() {
        return Objects.hash(elements, diagonalProperties);
    }

    @Override
    public boolean equals(Object obj) {
        boolean result;
        if (obj != null && obj.getClass() == this.getClass()) {
            Border other = (Border) obj;
            result = elements.equals(other.elements) && diagonalProperties.equals(other.diagonalProperties);
        } else {
            result = false;
        }
        return result;
    }

    /**
     * Write this border definition as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<border");
        if (diagonalProperties.contains(DiagonalProperty.DIAGONAL_UP)) {
            w.append(" diagonalUp=\"1\"");
        }
        if (diagonalProperties.contains(DiagonalProperty.DIAGONAL_DOWN)) {
            w.append(" diagonalDown=\"1\"");
        }
        w.append(">");
        elements.get(BorderSide.LEFT).write("left", w);
        elements.get(BorderSide.RIGHT).write("right", w);
        elements.get(BorderSide.TOP).write("top", w);
        elements.get(BorderSide.BOTTOM).write("bottom", w);
        elements.get(BorderSide.DIAGONAL).write("diagonal", w);
        w.append("</border>");
    }

}
