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
 * Represents an image embedded in a worksheet.
 */
public class Picture {

    private final int id;
    private final String name;
    private final PictureAnchor anchor;
    private final byte[] imageData;
    private final ImageType imageType;
    private final boolean lockAspectRatio;

    // Relationship ID (set when writing)
    private String relationshipId;

    Picture(int id, String name, PictureAnchor anchor, byte[] imageData, ImageType imageType,
            boolean lockAspectRatio) {
        this.id = id;
        this.name = name != null ? name : "Picture " + id;
        this.anchor = anchor;
        this.imageData = imageData;
        this.imageType = imageType;
        this.lockAspectRatio = lockAspectRatio;
    }

    void setRelationshipId(String relationshipId) {
        this.relationshipId = relationshipId;
    }

    String getRelationshipId() {
        return relationshipId;
    }

    int getId() {
        return id;
    }

    byte[] getImageData() {
        return imageData;
    }

    ImageType getImageType() {
        return imageType;
    }

    PictureAnchor getAnchor() {
        return anchor;
    }

    /**
     * Get the name of this picture.
     *
     * @return Picture name
     */
    public String getName() {
        return name;
    }

    /**
     * Write the picture element to the drawing XML.
     */
    void write(Writer w) throws IOException {
        if (anchor.isTwoCellAnchor()) {
            writeTwoCellAnchor(w);
        } else {
            writeOneCellAnchor(w);
        }
    }

    private void writeOneCellAnchor(Writer w) throws IOException {
        w.append("<xdr:oneCellAnchor>");
        anchor.writeFrom(w);
        anchor.writeExt(w);
        writePicElement(w);
        w.append("<xdr:clientData/>");
        w.append("</xdr:oneCellAnchor>");
    }

    private void writeTwoCellAnchor(Writer w) throws IOException {
        w.append("<xdr:twoCellAnchor>");
        anchor.writeFrom(w);
        anchor.writeTo(w);
        writePicElement(w);
        w.append("<xdr:clientData/>");
        w.append("</xdr:twoCellAnchor>");
    }

    private void writePicElement(Writer w) throws IOException {
        w.append("<xdr:pic>");

        // Non-visual properties
        w.append("<xdr:nvPicPr>");
        w.append("<xdr:cNvPr id=\"").append(id).append("\" name=\"");
        w.appendEscaped(name);
        w.append("\"/>");
        w.append("<xdr:cNvPicPr>");
        if (lockAspectRatio) {
            w.append("<a:picLocks noChangeAspect=\"1\"/>");
        }
        w.append("</xdr:cNvPicPr>");
        w.append("</xdr:nvPicPr>");

        // Blip fill (image reference)
        w.append("<xdr:blipFill>");
        if (imageType == ImageType.SVG) {
            // SVG uses extension element with svgBlip (Office 2016+)
            w.append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            w.append("<a:extLst>");
            w.append("<a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">");
            w.append("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" ");
            w.append("r:embed=\"").append(relationshipId).append("\"/>");
            w.append("</a:ext>");
            w.append("</a:extLst>");
            w.append("</a:blip>");
        } else {
            // Raster images (PNG, JPEG, GIF)
            w.append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
            w.append("r:embed=\"").append(relationshipId).append("\"/>");
        }
        w.append("<a:stretch><a:fillRect/></a:stretch>");
        w.append("</xdr:blipFill>");

        // Shape properties
        w.append("<xdr:spPr>");
        w.append("<a:xfrm>");
        w.append("<a:off x=\"0\" y=\"0\"/>");
        if (anchor.isTwoCellAnchor()) {
            w.append("<a:ext cx=\"0\" cy=\"0\"/>");
        } else {
            w.append("<a:ext cx=\"").append(anchor.getWidthEmu()).append("\" cy=\"")
              .append(anchor.getHeightEmu()).append("\"/>");
        }
        w.append("</a:xfrm>");
        w.append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
        w.append("</xdr:spPr>");

        w.append("</xdr:pic>");
    }
}
