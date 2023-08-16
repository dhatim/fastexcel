package org.dhatim.fastexcel;

import java.io.IOException;

public class Table{
    int index;
    private String name ;
    private String displayName;
    private boolean totalsRowShown = false;
    private Range range;
    private String[] headers ;

    private final TableStyleInfo styleInfo = new TableStyleInfo(this);

    Table(int index, Range range, String[] headers) {
        int count = range.getRight() - range.getLeft() + 1;
        if (headers.length != count) {
            throw new IllegalStateException("Header length no match the count of columns,table index:" + index);
        }
        for (int i = 0; i < count; i++) {
            range.getWorksheet().value(range.getTop(), range.getLeft() + i, headers[i]);
        }
        this.index = index;
        this.range = range;
        this.headers = headers;
    }

    public Table setName(String name) {
        this.name = name;
        return this;
    }

    public Table setDisplayName(String displayName) {
        this.displayName = displayName;
        return this;
    }

    public Table setTotalsRowShown(boolean totalsRowShown) {
        this.totalsRowShown = totalsRowShown;
        return this;
    }

    public TableStyleInfo styleInfo() {
        return styleInfo;
    }

    public class TableStyleInfo {
        private final Table table;
        TableStyleInfo(Table table){
            this.table = table;
        }
        private String name ;
        private boolean showFirstColumn = false;
        private boolean showLastColumn = false;
        private boolean showRowStripes = true;
        private boolean showColumnStripes = false;

        public TableStyleInfo setStyleName(String name) {
            this.name = name;
            return this;
        }

        public TableStyleInfo setShowFirstColumn(boolean showFirstColumn) {
            this.showFirstColumn = showFirstColumn;
            return this;
        }

        public TableStyleInfo setShowLastColumn(boolean showLastColumn) {
            this.showLastColumn = showLastColumn;
            return this;
        }

        public TableStyleInfo setShowRowStripes(boolean showRowStripes) {
            this.showRowStripes = showRowStripes;
            return this;
        }

        public TableStyleInfo setShowColumnStripes(boolean showColumnStripes) {
            this.showColumnStripes = showColumnStripes;
            return this;
        }

        public void write(Writer w) throws IOException {
            w.append("<tableStyleInfo name=\""+(name==null||"".equals(name)?"TableStyleMedium2":styleInfo.name)+"\" ");
            w.append("showFirstColumn=\""+(showFirstColumn?1:0)+"\" ");
            w.append("showLastColumn=\""+(showLastColumn?1:0)+"\" ");
            w.append("showRowStripes=\""+(showRowStripes?1:0)+"\" ");
            w.append("showColumnStripes=\""+(showColumnStripes?1:0)+"\"/>");
            w.append("</table>");
        }
    }

    void write(Writer w) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        w.append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
        w.append("id=\"" + index + "\" ");
        w.append("name=\""+(name==null||"".equals(name)?"Table"+index:name)+"\" ");
        w.append("displayName=\"" + (displayName==null||"".equals(displayName)?"Table"+index:displayName) + "\" ");
        w.append("ref=\"" + range.toString() + "\" ");
        w.append("totalsRowShown=\"" + (totalsRowShown ? 1 : 0) + "\">");
        w.append("<autoFilter ref=\"" + range.toString() + "\"/>");
        int count = range.getRight() - range.getLeft() + 1;
        w.append("<tableColumns count=\"" + count + "\">");
        for (int i = 0; i < count; i++) {
            w.append("<tableColumn id=\"" + (i + 1) + "\" name=\"" + headers[i] + "\"/>");
        }
        w.append("</tableColumns>");
        styleInfo.write(w);
    }

}
