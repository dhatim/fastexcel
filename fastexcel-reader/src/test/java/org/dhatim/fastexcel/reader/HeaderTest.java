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
package org.dhatim.fastexcel.reader;

import static org.assertj.core.api.Assertions.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;
import org.junit.jupiter.api.Test;

public class HeaderTest {

    @Test
    public void testInsufficientBytes() {
        byte[] bytes = new byte[1];
        assertThatThrownBy(() -> ReadableWorkbook.isOLE2Header(bytes)).isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> ReadableWorkbook.isOOXMLZipHeader(bytes)).isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    public void testOLE2File() throws IOException {
        byte[] bytes = readFirstBytes("/xls/defined_names_simple.xls", 8);
        assertThat(ReadableWorkbook.isOLE2Header(bytes)).isTrue();
        assertThat(ReadableWorkbook.isOOXMLZipHeader(bytes)).isFalse();
    }

    @Test
    public void testOOXMLZipFile() throws IOException {
        byte[] bytes = readFirstBytes("/xlsx/defined_names_simple.xlsx", 8);
        assertThat(ReadableWorkbook.isOLE2Header(bytes)).isFalse();
        assertThat(ReadableWorkbook.isOOXMLZipHeader(bytes)).isTrue();
    }

    private static byte[] readFirstBytes(String name, int length) throws IOException {
        try (InputStream is = open(name)) {
            byte[] bytes = new byte[8];
            readNBytes(is, bytes, 0, 8);
            return bytes;
        }
    }

    private static int readNBytes(InputStream is, byte[] b, int off, int len) throws IOException {
        if (off < 0 || len < 0 || len > b.length - off)
            throw new IndexOutOfBoundsException();
        int n = 0;
        while (n < len) {
            int count = is.read(b, off + n, len - n);
            if (count < 0)
                break;
            n += count;
        }
        return n;
    }

    private static InputStream open(String name) {
        InputStream result = HeaderTest.class.getResourceAsStream(name);
        if (result == null) {
            throw new IllegalStateException("Cannot read resource " + name);
        }
        return result;
    }

}
