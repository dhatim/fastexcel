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
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.nio.CharBuffer;
import java.nio.charset.CharsetEncoder;
import java.nio.charset.StandardCharsets;

/**
 * Writer for XML files.
 */
class Writer {

    /**
     * Target output stream.
     */
    private final OutputStream os;
    /**
     * Char buffer.
     */
    private final StringBuilder sb;
    /**
     * UTF-8 encoder for flush; reused to avoid allocating String and byte[] on every flush.
     */
    private final CharsetEncoder encoder = StandardCharsets.UTF_8.newEncoder();
    /**
     * Reusable char array for getChars (avoids sb.toString() in flush).
     */
    private char[] encodeChars;
    /**
     * Reusable byte array for encoded output (avoids getBytes() allocation in flush).
     */
    private byte[] encodeBytes;

    /**
     * Constructor.
     *
     * @param os Output stream.
     */
    Writer(OutputStream os) {
        this.os = os;
        this.sb = new StringBuilder(512 * 1024);
    }

    /**
     * Append a string without escaping.
     *
     * @param s String.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(String s) throws IOException {
        return append(s, false);
    }

    /**
     * Append a string with XML escaping.
     *
     * @param s String.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer appendEscaped(String s) throws IOException {
        return append(s, true);
    }

    /**
     * Append a string with or without escaping.
     *
     * @param s String.
     * @param escape Whether the string should be escaped or not
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    private Writer append(String s, boolean escape) throws IOException {
        if (escape) {
            XmlEscapeHelper.appendEscaped(sb, s);
        } else {
            sb.append(s);
        }
        check();
        return this;
    }

    /**
     * Check if the buffer gets full. In this case, flush bytes to the output
     * stream.
     *
     * @throws IOException If an I/O error occurs.
     */
    private void check() throws IOException {
        if (sb.capacity() - sb.length() < 1024) {
            flush();
        }
    }

    /**
     * Append a char without escaping.
     *
     * @param c Character.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(char c) throws IOException {
        sb.append(c);
        check();
        return this;
    }

    /**
     * Append an integer.
     *
     * @param n Integer.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(int n) throws IOException {
        sb.append(n);
        check();
        return this;
    }

    /**
     * Append a long.
     *
     * @param n Long.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(long n) throws IOException {
        sb.append(n);
        check();
        return this;
    }

    /**
     * Append a double.
     *
     * @param n Double.
     * @return This writer.
     * @throws IOException If an I/O error occurs.
     */
    Writer append(double n) throws IOException {
        sb.append(n);
        check();
        return this;
    }

    /**
     * Flush this writer.
     * Encodes buffered characters to UTF-8 and writes to the output stream
     * without allocating a full String and byte[] (avoids double copy).
     *
     * @throws IOException If an I/O error occurs.
     */
    void flush() throws IOException {
        int len = sb.length();
        if (len == 0) {
            return;
        }
        if (encodeChars == null || encodeChars.length < len) {
            encodeChars = new char[Math.max(len, 512 * 1024)];
        }
        sb.getChars(0, len, encodeChars, 0);
        sb.setLength(0);

        int estimatedBytes = len * 4; // UTF-8 max bytes per code point
        if (encodeBytes == null || encodeBytes.length < estimatedBytes) {
            encodeBytes = new byte[Math.max(estimatedBytes, 512 * 1024 * 4)];
        }
        ByteBuffer out = ByteBuffer.wrap(encodeBytes);
        encoder.reset();
        encoder.encode(CharBuffer.wrap(encodeChars, 0, len), out, true);
        encoder.flush(out);
        os.write(encodeBytes, 0, out.position());
    }
}
