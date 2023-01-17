package org.dhatim.fastexcel;

public class HyperLink {
    private final String displayStr;

    private final String linkStr;

    public HyperLink(String linkStr) {
        this.linkStr = linkStr;
        this.displayStr = linkStr;
    }

    public HyperLink(String linkStr, String displayStr) {
        this.linkStr = linkStr;
        this.displayStr = displayStr;
    }

    @Override
    public int hashCode() {
        return super.hashCode();
    }

    @Override
    public boolean equals(Object obj) {
        return super.equals(obj);
    }

    public String getDisplayStr() {
        return displayStr;
    }

    public String getLinkStr() {
        return linkStr;
    }
}
