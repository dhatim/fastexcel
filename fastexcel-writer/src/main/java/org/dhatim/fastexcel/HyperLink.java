package org.dhatim.fastexcel;

import java.util.Objects;

public class HyperLink {
    private final String displayStr;

    private final String linkStr;

    private final HyperLinkType hyperLinkType;

    /**
     * Static factory method which allows to create external HyperLink
     * @param linkStr external link for which the hyperlink will lead to
     * @param displayStr string which will displayed in hyperlink cell
     * @return External HyperLink
     */
    public static HyperLink external(String linkStr, String displayStr) {
        return new HyperLink(linkStr, displayStr, HyperLinkType.EXTERNAL);
    }

    /**
     * Static factory method which allows to create internal HyperLink
     * @param linkStr link for which the hyperlink will lead to
     * @param displayStr string which will displayed in hyperlink cell
     * @return Internal HyperLink
     */
    public static HyperLink internal(String linkStr, String displayStr) {
        return new HyperLink(linkStr, displayStr, HyperLinkType.INTERNAL);
    }

    public HyperLink(String linkStr) {
        this(linkStr, linkStr, HyperLinkType.EXTERNAL);
    }

    /**
     * Default Constructor
     * By default, the HyperLink will be marked as an external
     * @param linkStr external link for which the hyperlink will lead to
     * @param displayStr string which will displayed in hyperlink cell
     */
    public HyperLink(String linkStr, String displayStr) {
        this(linkStr, displayStr, HyperLinkType.EXTERNAL);
    }

    /**
     * Constructor
     * @param linkStr link for which the hyperlink will lead to
     * @param displayStr string which will displayed in hyperlink cell
     * @param hyperLinkType identifies type of the hyperlink
     */
    HyperLink(String linkStr, String displayStr, HyperLinkType hyperLinkType) {
        this.linkStr = linkStr;
        this.displayStr = displayStr;
        this.hyperLinkType = hyperLinkType;
    }

    @Override
    public boolean equals(Object o) {
        if (o == null || getClass() != o.getClass()) return false;
        HyperLink hyperLink = (HyperLink) o;
        return Objects.equals(displayStr, hyperLink.displayStr)
                && Objects.equals(linkStr, hyperLink.linkStr)
                && hyperLinkType == hyperLink.hyperLinkType;
    }

    @Override
    public int hashCode() {
        return Objects.hash(displayStr, linkStr, hyperLinkType);
    }

    @Override
    public String toString() {
        return linkStr;
    }

    public String getDisplayStr() {
        return displayStr;
    }

    public String getLinkStr() {
        return linkStr;
    }

    HyperLinkType getHyperLinkType() {
        return hyperLinkType;
    }
}
