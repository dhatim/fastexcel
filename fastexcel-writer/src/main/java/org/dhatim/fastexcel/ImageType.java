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
 * Supported image types for embedding in worksheets.
 */
public enum ImageType {
    PNG("png", "image/png", false),
    JPEG("jpeg", "image/jpeg", false),
    GIF("gif", "image/gif", false),
    SVG("svg", "image/svg+xml", true);

    private final String extension;
    private final String contentType;
    private final boolean vector;

    ImageType(String extension, String contentType, boolean vector) {
        this.extension = extension;
        this.contentType = contentType;
        this.vector = vector;
    }

    /**
     * Check if this is a vector image format (e.g., SVG).
     *
     * @return true if vector format, false if raster
     */
    public boolean isVector() {
        return vector;
    }

    /**
     * Get the file extension for this image type.
     *
     * @return File extension without the dot (e.g., "png", "jpeg")
     */
    public String getExtension() {
        return extension;
    }

    /**
     * Get the MIME content type for this image type.
     *
     * @return MIME content type (e.g., "image/png")
     */
    public String getContentType() {
        return contentType;
    }

    /**
     * Detect image type from byte array header.
     *
     * @param data Image bytes
     * @return Detected ImageType
     * @throws IllegalArgumentException if the image format is not supported or data is invalid
     */
    public static ImageType fromBytes(byte[] data) {
        if (data == null || data.length < 8) {
            throw new IllegalArgumentException("Invalid image data: data is null or too short");
        }
        // PNG signature: 89 50 4E 47 0D 0A 1A 0A
        if (data[0] == (byte) 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
                && data[4] == 0x0D && data[5] == 0x0A && data[6] == 0x1A && data[7] == 0x0A) {
            return PNG;
        }
        // JPEG signature: FF D8 FF
        if (data[0] == (byte) 0xFF && data[1] == (byte) 0xD8 && data[2] == (byte) 0xFF) {
            return JPEG;
        }
        // GIF signature: 47 49 46 38 (GIF8)
        if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38) {
            return GIF;
        }
        // SVG detection: look for <?xml or <svg in the beginning (text-based format)
        String header = new String(data, 0, Math.min(data.length, 256), java.nio.charset.StandardCharsets.UTF_8);
        if (header.contains("<svg") || (header.contains("<?xml") && header.contains("<svg"))) {
            return SVG;
        }
        throw new IllegalArgumentException("Unsupported image format. Supported formats: PNG, JPEG, GIF, SVG");
    }
}
