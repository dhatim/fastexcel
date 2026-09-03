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

/**
 * Defines the positioning of an image within a worksheet.
 * Supports both one-cell and two-cell anchoring.
 * <p>
 * One-cell anchor: Image is positioned at a cell with explicit size (width/height).
 * Two-cell anchor: Image spans from one cell to another, resizing with cells.
 */
public class PictureAnchor {

    /**
     * EMUs (English Metric Units) per pixel at 96 DPI.
     */
    public static final int EMU_PER_PIXEL = 9525;

    /**
     * EMUs per point.
     */
    public static final int EMU_PER_POINT = 12700;

    /**
     * EMUs per inch.
     */
    public static final int EMU_PER_INCH = 914400;

    /**
     * EMUs per centimeter.
     */
    public static final int EMU_PER_CM = 360000;

    private final AnchorGeometry geometry;

    /**
     * Create a one-cell anchor with explicit size in pixels.
     *
     * @param row      Zero-based row number
     * @param col      Zero-based column number
     * @param widthPx  Image width in pixels
     * @param heightPx Image height in pixels
     * @return A new PictureAnchor configured for one-cell anchoring
     */
    public static PictureAnchor oneCellAnchor(int row, int col, int widthPx, int heightPx) {
        return new PictureAnchor(new OneCellAnchorGeometry(new AnchorMarker(row, col, 0, 0),
                new AnchorExtent((long) widthPx * EMU_PER_PIXEL, (long) heightPx * EMU_PER_PIXEL)));
    }

    /**
     * Create a one-cell anchor with offset and explicit size in pixels.
     *
     * @param row       Zero-based row number
     * @param col       Zero-based column number
     * @param colOffPx  Column offset in pixels from the left edge of the cell
     * @param rowOffPx  Row offset in pixels from the top edge of the cell
     * @param widthPx   Image width in pixels
     * @param heightPx  Image height in pixels
     * @return A new PictureAnchor configured for one-cell anchoring with offset
     */
    public static PictureAnchor oneCellAnchor(int row, int col, int colOffPx, int rowOffPx,
                                               int widthPx, int heightPx) {
        return new PictureAnchor(new OneCellAnchorGeometry(new AnchorMarker(row, col,
                colOffPx * EMU_PER_PIXEL, rowOffPx * EMU_PER_PIXEL),
                new AnchorExtent((long) widthPx * EMU_PER_PIXEL, (long) heightPx * EMU_PER_PIXEL)));
    }

    /**
     * Create a two-cell anchor spanning from one cell to another.
     *
     * @param fromRow Starting row (zero-based)
     * @param fromCol Starting column (zero-based)
     * @param toRow   Ending row (zero-based, exclusive - image ends at top of this row)
     * @param toCol   Ending column (zero-based, exclusive - image ends at left of this column)
     * @return A new PictureAnchor configured for two-cell anchoring
     */
    public static PictureAnchor twoCellAnchor(int fromRow, int fromCol, int toRow, int toCol) {
        return new PictureAnchor(new TwoCellAnchorGeometry(new AnchorMarker(fromRow, fromCol, 0, 0),
                new AnchorMarker(toRow, toCol, 0, 0)));
    }

    /**
     * Create a two-cell anchor with offsets.
     *
     * @param fromRow      Starting row (zero-based)
     * @param fromCol      Starting column (zero-based)
     * @param fromColOffPx Starting column offset in pixels
     * @param fromRowOffPx Starting row offset in pixels
     * @param toRow        Ending row (zero-based)
     * @param toCol        Ending column (zero-based)
     * @param toColOffPx   Ending column offset in pixels
     * @param toRowOffPx   Ending row offset in pixels
     * @return A new PictureAnchor configured for two-cell anchoring with offsets
     */
    public static PictureAnchor twoCellAnchor(int fromRow, int fromCol, int fromColOffPx, int fromRowOffPx,
                                               int toRow, int toCol, int toColOffPx, int toRowOffPx) {
        return new PictureAnchor(new TwoCellAnchorGeometry(new AnchorMarker(fromRow, fromCol,
                fromColOffPx * EMU_PER_PIXEL, fromRowOffPx * EMU_PER_PIXEL),
                new AnchorMarker(toRow, toCol, toColOffPx * EMU_PER_PIXEL, toRowOffPx * EMU_PER_PIXEL)));
    }

    private PictureAnchor(AnchorGeometry geometry) {
        this.geometry = geometry;
    }

    /**
     * Check if this anchor is a two-cell anchor.
     *
     * @return true if two-cell anchor, false if one-cell anchor
     */
    public boolean isTwoCellAnchor() {
        return geometry.isTwoCellAnchor();
    }

    /**
     * Write the "from" position element.
     */
    void writeFrom(Writer w) throws IOException {
        geometry.writeFrom(w);
    }

    /**
     * Write the "to" position element (for two-cell anchors).
     */
    void writeTo(Writer w) throws IOException {
        geometry.writeTo(w);
    }

    /**
     * Write the extent element (for one-cell anchors).
     */
    void writeExt(Writer w) throws IOException {
        geometry.writeExt(w);
    }

    /**
     * Get width in EMUs (for one-cell anchors).
     */
    public Long getWidthEmu() {
        return geometry.getWidthEmu();
    }

    /**
     * Get height in EMUs (for one-cell anchors).
     */
    public Long getHeightEmu() {
        return geometry.getHeightEmu();
    }

    private interface AnchorGeometry {
        boolean isTwoCellAnchor();

        void writeFrom(Writer w) throws IOException;

        void writeTo(Writer w) throws IOException;

        void writeExt(Writer w) throws IOException;

        Long getWidthEmu();

        Long getHeightEmu();
    }

    private static final class OneCellAnchorGeometry implements AnchorGeometry {
        private final AnchorMarker from;
        private final AnchorExtent extent;

        private OneCellAnchorGeometry(AnchorMarker from, AnchorExtent extent) {
            this.from = from;
            this.extent = extent;
        }

        public boolean isTwoCellAnchor() {
            return false;
        }

        public void writeFrom(Writer w) throws IOException {
            from.write(w, "from");
        }

        public void writeTo(Writer w) {
            throw new UnsupportedOperationException("One-cell anchors do not have a to marker.");
        }

        public void writeExt(Writer w) throws IOException {
            extent.write(w);
        }

        public Long getWidthEmu() {
            return extent.widthEmu;
        }

        public Long getHeightEmu() {
            return extent.heightEmu;
        }
    }

    private static final class TwoCellAnchorGeometry implements AnchorGeometry {
        private final AnchorMarker from;
        private final AnchorMarker to;

        private TwoCellAnchorGeometry(AnchorMarker from, AnchorMarker to) {
            this.from = from;
            this.to = to;
        }

        public boolean isTwoCellAnchor() {
            return true;
        }

        public void writeFrom(Writer w) throws IOException {
            from.write(w, "from");
        }

        public void writeTo(Writer w) throws IOException {
            to.write(w, "to");
        }

        public void writeExt(Writer w) {
            throw new UnsupportedOperationException("Two-cell anchors do not have an extent.");
        }

        public Long getWidthEmu() {
            return null;
        }

        public Long getHeightEmu() {
            return null;
        }
    }

    private static final class AnchorMarker {
        private final int row;
        private final int col;
        private final int colOffsetEmu;
        private final int rowOffsetEmu;

        private AnchorMarker(int row, int col, int colOffsetEmu, int rowOffsetEmu) {
            this.row = row;
            this.col = col;
            this.colOffsetEmu = colOffsetEmu;
            this.rowOffsetEmu = rowOffsetEmu;
        }

        private void write(Writer w, String elementName) throws IOException {
            w.append("<xdr:").append(elementName).append(">");
            w.append("<xdr:col>").append(col).append("</xdr:col>");
            w.append("<xdr:colOff>").append(colOffsetEmu).append("</xdr:colOff>");
            w.append("<xdr:row>").append(row).append("</xdr:row>");
            w.append("<xdr:rowOff>").append(rowOffsetEmu).append("</xdr:rowOff>");
            w.append("</xdr:").append(elementName).append(">");
        }
    }

    private static final class AnchorExtent {
        private final long widthEmu;
        private final long heightEmu;

        private AnchorExtent(long widthEmu, long heightEmu) {
            this.widthEmu = widthEmu;
            this.heightEmu = heightEmu;
        }

        private void write(Writer w) throws IOException {
            w.append("<xdr:ext cx=\"").append(widthEmu).append("\" cy=\"").append(heightEmu).append("\"/>");
        }
    }
}
