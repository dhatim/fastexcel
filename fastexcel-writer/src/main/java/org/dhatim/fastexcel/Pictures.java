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
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Manages all pictures in a worksheet.
 */
class Pictures {

    private final List<Picture> pictures = new ArrayList<>();
    private final AtomicInteger idCounter = new AtomicInteger(1);
    private final Set<ImageType> usedImageTypes = new HashSet<>();

    /**
     * Add a picture with one-cell anchor.
     *
     * @param row       Zero-based row number
     * @param col       Zero-based column number
     * @param imageData Image bytes
     * @param widthPx   Width in pixels
     * @param heightPx  Height in pixels
     * @return The created Picture
     */
    Picture addPicture(int row, int col, byte[] imageData, int widthPx, int heightPx) {
        return addPicture(PictureAnchor.oneCellAnchor(row, col, widthPx, heightPx),
                imageData, null, true);
    }

    /**
     * Add a picture with two-cell anchor.
     *
     * @param fromRow   Starting row (zero-based)
     * @param fromCol   Starting column (zero-based)
     * @param toRow     Ending row (zero-based)
     * @param toCol     Ending column (zero-based)
     * @param imageData Image bytes
     * @return The created Picture
     */
    Picture addPicture(int fromRow, int fromCol, int toRow, int toCol, byte[] imageData) {
        return addPicture(PictureAnchor.twoCellAnchor(fromRow, fromCol, toRow, toCol),
                imageData, null, true);
    }

    /**
     * Add a picture with custom anchor.
     *
     * @param anchor          The positioning anchor
     * @param imageData       Image bytes
     * @param name            Picture name (can be null)
     * @param lockAspectRatio Whether to lock aspect ratio
     * @return The created Picture
     */
    Picture addPicture(PictureAnchor anchor, byte[] imageData, String name, boolean lockAspectRatio) {
        ImageType imageType = ImageType.fromBytes(imageData);
        usedImageTypes.add(imageType);

        int id = idCounter.getAndIncrement();
        Picture picture = new Picture(id, name, anchor, imageData, imageType, lockAspectRatio);
        pictures.add(picture);
        return picture;
    }

    boolean isEmpty() {
        return pictures.isEmpty();
    }

    int size() {
        return pictures.size();
    }

    Set<ImageType> getUsedImageTypes() {
        return Collections.unmodifiableSet(usedImageTypes);
    }

    List<Picture> getPictures() {
        return Collections.unmodifiableList(pictures);
    }

    /**
     * Write the drawing XML file.
     */
    void writeDrawing(Writer w) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        w.append("<xdr:wsDr ");
        w.append("xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ");
        w.append("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");

        for (Picture picture : pictures) {
            picture.write(w);
        }

        w.append("</xdr:wsDr>");
    }

    /**
     * Write the drawing relationships file.
     *
     * @param w          The writer
     * @param sheetIndex The sheet index (1-based)
     */
    void writeDrawingRels(Writer w, int sheetIndex) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        w.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");

        int imageIndex = 1;
        for (Picture picture : pictures) {
            String rId = "rId" + imageIndex;
            picture.setRelationshipId(rId);
            String imageName = "image" + sheetIndex + "_" + imageIndex + "." + picture.getImageType().getExtension();

            w.append("<Relationship Id=\"").append(rId).append("\" ");
            w.append("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" ");
            w.append("Target=\"../media/").append(imageName).append("\"/>");
            imageIndex++;
        }

        w.append("</Relationships>");
    }

    /**
     * Write image files to the media folder.
     *
     * @param workbook   The parent workbook
     * @param sheetIndex The sheet index (1-based)
     */
    void writeMediaFiles(Workbook workbook, int sheetIndex) throws IOException {
        int imageIndex = 1;
        for (Picture picture : pictures) {
            String imageName = "image" + sheetIndex + "_" + imageIndex + "." + picture.getImageType().getExtension();
            workbook.writeBinaryFile("xl/media/" + imageName, picture.getImageData());
            imageIndex++;
        }
    }
}
