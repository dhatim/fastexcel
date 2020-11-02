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

/** This class represents either a range of columns or a range of rows
 *  for the repeating cols/rows in the print setup. */
public class RepeatRange { 
    
    final int from;
    final int to;

    public RepeatRange(int from, int to) {
        this.from = from;
        this.to = to;
    }

    /**
     * Column indexes need to be transformed to the letter form.
     */
    public String colRangeToString() {
        return "$" + Range.colToString(from) + ":$" + Range.colToString(to);
    }

    /**
     * Row indexes need to be increased by 1 
     * (sheet rows start from 1 and not from 0)
     */
    public String rowRangeToString() {
        return "$" + (1 + from) + ":$" + (1 + to);
    }
}