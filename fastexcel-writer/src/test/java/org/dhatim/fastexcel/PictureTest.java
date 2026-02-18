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

import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class PictureTest {

    /**
     * Create a minimal valid 1x1 red PNG for testing.
     * This is a complete PNG file that can be read by image parsers.
     */
    private static byte[] createTestPng() {
        return new byte[] {
            // PNG signature
            (byte) 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            // IHDR chunk (13 bytes data)
            0x00, 0x00, 0x00, 0x0D, // length
            0x49, 0x48, 0x44, 0x52, // "IHDR"
            0x00, 0x00, 0x00, 0x01, // width = 1
            0x00, 0x00, 0x00, 0x01, // height = 1
            0x08, // bit depth = 8
            0x02, // color type = RGB
            0x00, // compression method
            0x00, // filter method
            0x00, // interlace method
            (byte) 0x90, 0x77, 0x53, (byte) 0xDE, // CRC
            // IDAT chunk (compressed image data)
            0x00, 0x00, 0x00, 0x0C, // length
            0x49, 0x44, 0x41, 0x54, // "IDAT"
            0x08, (byte) 0xD7, // zlib header
            0x63, (byte) 0xF8, (byte) 0xCF, (byte) 0xC0, 0x00, 0x00, // compressed data
            0x00, 0x03, 0x00, 0x01, // more compressed data
            0x00, 0x05, // checksum part
            (byte) 0xFE, (byte) 0xD4, (byte) 0xEF, (byte) 0xA5, // CRC
            // IEND chunk
            0x00, 0x00, 0x00, 0x00, // length = 0
            0x49, 0x45, 0x4E, 0x44, // "IEND"
            (byte) 0xAE, 0x42, 0x60, (byte) 0x82 // CRC
        };
    }

    /**
     * Create a minimal valid JPEG for testing.
     */
    private static byte[] createTestJpeg() {
        return new byte[] {
            // JPEG signature (SOI marker)
            (byte) 0xFF, (byte) 0xD8,
            // APP0 marker (JFIF)
            (byte) 0xFF, (byte) 0xE0,
            0x00, 0x10, // length
            0x4A, 0x46, 0x49, 0x46, 0x00, // "JFIF\0"
            0x01, 0x01, // version
            0x00, // units
            0x00, 0x01, // X density
            0x00, 0x01, // Y density
            0x00, 0x00, // thumbnail size
            // SOF0 marker (start of frame)
            (byte) 0xFF, (byte) 0xC0,
            0x00, 0x0B, // length
            0x08, // precision
            0x00, 0x01, // height = 1
            0x00, 0x01, // width = 1
            0x01, // components
            0x01, 0x11, 0x00, // component info
            // DHT marker (Huffman table)
            (byte) 0xFF, (byte) 0xC4,
            0x00, 0x14, // length
            0x00, // table info
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00,
            // SOS marker (start of scan)
            (byte) 0xFF, (byte) 0xDA,
            0x00, 0x08, // length
            0x01, // components
            0x01, 0x00, // component selector
            0x00, 0x3F, 0x00, // spectral selection
            // Image data
            0x7F,
            // EOI marker (end of image)
            (byte) 0xFF, (byte) 0xD9
        };
    }

    @Test
    void testImageTypeDetectionPng() {
        byte[] png = createTestPng();
        assertThat(ImageType.fromBytes(png)).isEqualTo(ImageType.PNG);
    }

    @Test
    void testImageTypeDetectionJpeg() {
        byte[] jpeg = createTestJpeg();
        assertThat(ImageType.fromBytes(jpeg)).isEqualTo(ImageType.JPEG);
    }

    @Test
    void testImageTypeDetectionGif() {
        byte[] gif = new byte[] {0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x00, 0x00};
        assertThat(ImageType.fromBytes(gif)).isEqualTo(ImageType.GIF);
    }

    @Test
    void testImageTypeDetectionSvg() {
        String svg = "<?xml version=\"1.0\"?><svg xmlns=\"http://www.w3.org/2000/svg\"><circle/></svg>";
        byte[] svgBytes = svg.getBytes(java.nio.charset.StandardCharsets.UTF_8);
        assertThat(ImageType.fromBytes(svgBytes)).isEqualTo(ImageType.SVG);
    }

    @Test
    void testImageTypeDetectionSvgWithoutXmlDeclaration() {
        String svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100\" height=\"100\"><rect/></svg>";
        byte[] svgBytes = svg.getBytes(java.nio.charset.StandardCharsets.UTF_8);
        assertThat(ImageType.fromBytes(svgBytes)).isEqualTo(ImageType.SVG);
    }

    @Test
    void testImageTypeDetectionUnsupported() {
        byte[] unknown = new byte[] {0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07};
        assertThatThrownBy(() -> ImageType.fromBytes(unknown))
            .isInstanceOf(IllegalArgumentException.class)
            .hasMessageContaining("Unsupported image format");
    }

    @Test
    void testImageTypeDetectionNull() {
        assertThatThrownBy(() -> ImageType.fromBytes(null))
            .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void testImageTypeDetectionTooShort() {
        byte[] tooShort = new byte[] {0x00, 0x01, 0x02};
        assertThatThrownBy(() -> ImageType.fromBytes(tooShort))
            .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void testAddImageOneCellAnchor() throws Exception {
        byte[] imageData = createTestPng();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, "Logo:");
            ws.addImage(0, 1, imageData, 100, 50);
        }

        // Verify with Apache POI
        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            XSSFSheet sheet = poiWb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertThat(drawing).isNotNull();
            assertThat(drawing.getShapes()).hasSize(1);
            assertThat(drawing.getShapes().get(0)).isInstanceOf(XSSFPicture.class);
        }
    }

    @Test
    void testAddImageTwoCellAnchor() throws Exception {
        byte[] imageData = createTestPng();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.addImage(0, 0, 5, 3, imageData);
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            XSSFSheet sheet = poiWb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertThat(drawing).isNotNull();
            assertThat(drawing.getShapes()).hasSize(1);
        }
    }

    @Test
    void testMultipleImages() throws Exception {
        byte[] imageData = createTestPng();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.addImage(0, 0, imageData, 50, 50);
            ws.addImage(1, 0, imageData, 50, 50);
            ws.addImage(2, 0, imageData, 50, 50);
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            XSSFSheet sheet = poiWb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertThat(drawing.getShapes()).hasSize(3);
        }
    }

    @Test
    void testImageWithComments() throws Exception {
        byte[] imageData = createTestPng();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, "Cell with comment");
            ws.comment(0, 0, "This is a comment");
            ws.addImage(1, 0, imageData, 100, 100);
        }

        // Verify both comments and images work together
        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            XSSFSheet sheet = poiWb.getSheetAt(0);
            // Check comment exists
            assertThat(sheet.getCellComment(new org.apache.poi.ss.util.CellAddress(0, 0))).isNotNull();
            // Check drawing exists
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertThat(drawing).isNotNull();
        }
    }

    @Test
    void testMultipleWorksheets() throws Exception {
        byte[] imageData = createTestPng();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws1 = wb.newWorksheet("Sheet1");
            ws1.addImage(0, 0, imageData, 100, 100);

            Worksheet ws2 = wb.newWorksheet("Sheet2");
            ws2.addImage(0, 0, imageData, 150, 150);
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            assertThat(poiWb.getSheetAt(0).getDrawingPatriarch().getShapes()).hasSize(1);
            assertThat(poiWb.getSheetAt(1).getDrawingPatriarch().getShapes()).hasSize(1);
        }
    }

    @Test
    void testImageWithCustomAnchor() throws Exception {
        byte[] imageData = createTestPng();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            PictureAnchor anchor = PictureAnchor.oneCellAnchor(2, 1, 10, 10, 200, 150);
            ws.addImage(anchor, imageData, "MyImage", true);
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            XSSFSheet sheet = poiWb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertThat(drawing.getShapes()).hasSize(1);
        }
    }

    @Test
    void testPictureAnchorOneCellBasic() {
        PictureAnchor anchor = PictureAnchor.oneCellAnchor(5, 3, 100, 200);
        assertThat(anchor.isTwoCellAnchor()).isFalse();
        assertThat(anchor.getWidthEmu()).isEqualTo(100L * PictureAnchor.EMU_PER_PIXEL);
        assertThat(anchor.getHeightEmu()).isEqualTo(200L * PictureAnchor.EMU_PER_PIXEL);
    }

    @Test
    void testPictureAnchorTwoCellBasic() {
        PictureAnchor anchor = PictureAnchor.twoCellAnchor(0, 0, 5, 3);
        assertThat(anchor.isTwoCellAnchor()).isTrue();
        assertThat(anchor.getWidthEmu()).isNull();
        assertThat(anchor.getHeightEmu()).isNull();
    }

    @Test
    void testImageWithOtherContent() throws Exception {
        byte[] imageData = createTestPng();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            // Add regular content
            ws.value(0, 0, "Header");
            ws.value(1, 0, 123.45);
            ws.style(0, 0).bold().set();
            // Add merged cells
            ws.range(3, 0, 3, 2).merge();
            ws.value(3, 0, "Merged");
            // Add image
            ws.addImage(5, 0, imageData, 100, 100);
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            XSSFSheet sheet = poiWb.getSheetAt(0);
            // Verify content
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Header");
            assertThat(sheet.getRow(1).getCell(0).getNumericCellValue()).isEqualTo(123.45);
            // Verify merged region
            assertThat(sheet.getNumMergedRegions()).isEqualTo(1);
            // Verify image
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertThat(drawing.getShapes()).hasSize(1);
        }
    }

    @Test
    void testJpegImage() throws Exception {
        byte[] imageData = createTestJpeg();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = new Workbook(baos, "Test", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.addImage(0, 0, imageData, 100, 100);
        }

        try (XSSFWorkbook poiWb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            XSSFSheet sheet = poiWb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertThat(drawing.getShapes()).hasSize(1);
        }
    }
}
