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
 * Helper class to define a simple shading to alternate rows in a range of
 * cells.
 */
class AlternateShading extends Shading{

    /**
     * Constructor.
     *
     * @param range Range where alternate rows are shaded.
     * @param fill Index of cached fill pattern for shaded rows.
     */
    AlternateShading(Range range, int fill) {
        super(range, fill, 2);
    }
}
