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

/**
 * Helper class to define a simple shading to alternate rows in a range of
 * cells.
 */
class Shading {

    /**
     * Range where alternate rows are shaded.
     */
    private final Range range;
    /**
     *  Index of cached fill pattern for shaded rows.
     */
    private final int fill;
    /**
     * Shading frequency (each Nth row will be shaded).
     */
    private final int eachNRows;

    /**
     * Constructor.
     *
     * @param range Range where alternate rows are shaded.
     * @param fill Index of cached fill pattern for shaded rows.
     */
    
    Shading(Range range, int fill, int eachNRows) {
        this.range = range;
        this.fill = fill;
        this.eachNRows = eachNRows; 
    }

    /**
     * Get shading frequency.
     *
     * @return Shading row frequency.
     */
    public int getEachNRows() {
        return eachNRows;
    }

    /**
     * Write this definition of alternate shading as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        String formula = "NOT(MOD(ROW()-" + range.getTop() + "," + eachNRows + "))";
        w.append("<conditionalFormatting sqref=\"").append(range.toString()).append("\"><cfRule type=\"expression\" dxfId=\"").append(fill).append("\" priority=\"1\"><formula>" + formula + "</formula></cfRule></conditionalFormatting>");
    }
}
