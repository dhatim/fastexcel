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

import java.util.Comparator;
import java.util.Objects;

/**
 *  This class is used to refer to multiple rows.
 */
class RowRange implements Comparable<RowRange>, Ref {
    
    final int from;
    final int to;

    public RowRange(int from, int to) {
        this.from = from;
        this.to = to;
    }

    @Override
    public int compareTo(RowRange o) {
        return Comparator.comparingInt((RowRange rowRange) -> rowRange.from)
                .thenComparing((RowRange rowRange) -> rowRange.to)
                .compare(this, o);
    }

    @Override
    public boolean equals(Object obj) {
        if (obj instanceof RowRange) {
            RowRange that = (RowRange) obj;
            return this == that || (this.from == that.from && this.to == that.to);
        }
        return false;
    }

    @Override
    public int hashCode() {
        return Objects.hash(from,to);
    }

    /**
     * Row indexes need to be increased by 1
     *  (sheet row indexes start from 1 and not from 0)
     */
    @Override
    public String toString() {
        return "$" + (1 + from) + ":$" + (1 + to);
    }
}