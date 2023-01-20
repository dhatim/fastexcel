package org.dhatim.fastexcel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.atomic.AtomicInteger;

public class Relationships {

    private static final String TYPE_OF_HYPERLINK= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private static final String TYPE_OF_DRAWING= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    private static final String TYPE_OF_COMMENTS= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    private static final String TYPE_OF_VMLDRAWING= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    private static final String TYPE_OF_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

    private final AtomicInteger maxIndex = new AtomicInteger(1);

    final Worksheet worksheet;

    private ArrayList<Relationship> relationship = new ArrayList<>();

    public Relationships(Worksheet worksheet) {
        this.worksheet = worksheet;
    }

    String setHyperLinkRels(String target, String targetMode) {
        String id = "rId" + (maxIndex.getAndIncrement());
        relationship.add(new Relationship(id, TYPE_OF_HYPERLINK, target, targetMode));
        return id;
    }

    String setTableRels(int index) {
        String id = "rId" + (maxIndex.getAndIncrement());
        relationship.add(new Relationship(id,TYPE_OF_TABLE , "../tables/table"+index+".xml", null));
        return id;
    }

    void setCommentsRels(int index) {
        relationship.add(new Relationship("d", TYPE_OF_DRAWING, "../drawings/drawing" + index + ".xml", null));
        relationship.add(new Relationship("c", TYPE_OF_COMMENTS, "../comments" + index + ".xml", null));
        relationship.add(new Relationship("v", TYPE_OF_VMLDRAWING, "../drawings/vmlDrawing" + index + ".vml", null));
    }

    boolean isEmpty() {
        return relationship.isEmpty();
    }


    class Relationship {
        private String id ;
        private String type;

        private String target;

        private String targetMode;

        public Relationship(String id, String type, String target, String targetMode) {
            this.id = id;
            this.type = type;
            this.target = target;
            this.targetMode = targetMode;
        }
    }

    void write( Writer relsWr) throws IOException {
            relsWr.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>");
            relsWr.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (Relationship rs : relationship) {
                relsWr.append("<Relationship Id=\"" + rs.id + "\" ");
                relsWr.append("Target=\"" + rs.target + "\" ");
                if (rs.targetMode!=null) {
                    relsWr.append("TargetMode=\""+rs.targetMode+"\" " );
                }
                relsWr.append("Type=\""+rs.type+"\" ");
                relsWr.append( "/>");
            }
            relsWr.append("</Relationships>");
    }


}
