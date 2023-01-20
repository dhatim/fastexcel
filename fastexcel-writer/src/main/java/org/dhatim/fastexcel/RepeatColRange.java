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

/** This class represents a range of columns for the 
 *      repeating columns feature in the print setup. */
class RepeatColRange implements Ref {
    
    final int from;
    final int to;

    public RepeatColRange(int from, int to) {
        this.from = from;
        this.to = to;
    }

    /**
     * Column indexes need to be transformed to the letter form.
     */
    @Override
    public String toString() {
        return "$" + colToString(from) + ":$" + colToString(to);
    }
}