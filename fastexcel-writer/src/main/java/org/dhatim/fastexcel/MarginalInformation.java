package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * Marginal Information, represents a header or footer
 * in an Excel file.
 */
public class MarginalInformation {

	private static final String DEFAULT_FONT = "Times New Roman";
	private static final int DEFAULT_FONT_SIZE = 12;
	private final String text;
	private final Position position;
	private final String font;
	private final int fontSize;

	public MarginalInformation(String text, Position position) {
		this(text, position, DEFAULT_FONT, DEFAULT_FONT_SIZE);
	}

	private MarginalInformation(String text, Position position, String font, int fontSize) {
		this.text = text;
		this.position = position;
		this.font = font;
		this.fontSize = fontSize;
	}

	public MarginalInformation withFont(String font) {
		return new MarginalInformation(this.text, this.position, font, fontSize);
	}

	public MarginalInformation withFontSize(int fontSize) {
		return new MarginalInformation(this.text, this.position, this.font, fontSize);
	}

	public String getContent() {
		return "&amp;" + position.getPos() +
				"&amp;&quot;" + font + ",Regular&quot;&amp;" + fontSize +
				"&amp;K000000" + prepareTextForXml(text);
	}

	public void write(Writer writer) throws IOException {
		writer.append(getContent());
	}

	private String prepareTextForXml(String text) {
		switch (text.toLowerCase()) {
			case "page 1 of ?":
				return "Page &amp;P of &amp;N";
			case "page 1, sheetname":
				return "Page &amp;P, &amp;A";
			case "page 1":
				return "Page &amp;P";
			case "sheetname":
				return "&amp;A";
			default:
				XmlEscapeHelper xmlEscapeHelper = new XmlEscapeHelper();
				return xmlEscapeHelper.escape(text);
		}
	}
}
