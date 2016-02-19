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
 * Index of a cached string.
 */
public class CachedString {

    private final int index;

    /**
     * Constructor.
     *
     * @param index Index of cached string: return of value
     * {@link StringCache#cacheString(java.lang.String)}.
     */
    public CachedString(int index) {
        this.index = index;
    }

    /**
     * Get index of cached string.
     *
     * @return Index.
     */
    public int getIndex() {
        return index;
    }

    @Override
    public int hashCode() {
        int hash = 3;
        hash = 67 * hash + this.index;
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final CachedString other = (CachedString) obj;
        if (this.index != other.index) {
            return false;
        }
        return true;
    }

}
