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

import java.util.Objects;

/**
 * A cached string is uniquely identified by its index.
 */
class CachedString {

    private final String string;
    private final int index;

    /**
     * Constructor.
     *
     * @param string String value.
     * @param index Index in cache.
     */
    CachedString(String string, int index) {
        Objects.requireNonNull(string);
        this.string = string;
        this.index = index;
    }

    /**
     * Get string value.
     *
     * @return String value.
     */
    String getString() {
        return string;
    }

    /**
     * Get cache index.
     *
     * @return Cache index.
     */
    int getIndex() {
        return index;
    }

}
