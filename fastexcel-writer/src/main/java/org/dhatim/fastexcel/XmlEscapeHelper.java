package org.dhatim.fastexcel;

/**
 * Helper which can apply XML escaping to a string
 */
public class XmlEscapeHelper {

	/**
	 * Apply XML escaping to a String.
	 * Invalid characters in XML 1.0 are ignored.
	 *
	 * @param text text to be escaped
	 * @return escaped text
	 */
	public String escape(final String text) {
		int offset = 0;
		StringBuilder sb = new StringBuilder();
		while (offset < text.length()) {
			int codePoint = text.codePointAt(offset);
			sb.append(escape(codePoint));
			offset += Character.charCount(codePoint);
		}
		return sb.toString();
	}

	/**
	 * Escape char with XML escaping.
	 * Invalid characters in XML 1.0 are ignored.
	 *
	 * @param c Character code point.
	 */
	private String escape(int c) {
		if (!(c == 0x9 || c == 0xa || c == 0xD
				|| (c >= 0x20 && c <= 0xd7ff)
				|| (c >= 0xe000 && c <= 0xfffd)
				|| (c >= 0x10000 && c <= 0x10ffff))) {
			return "";
		}
		switch (c) {
			case '<':
				return "&lt;";
			case '>':
				return "&gt;";
			case '&':
				return "&amp;";
			case '\'':
				return "&apos;";
			case '"':
				return "&quot;";
			default:
				if (c > 0x7e || c < 0x20) {
					return "&#x".concat(Integer.toHexString(c)).concat(";");
				} else {
					return String.valueOf((char) c);
				}
		}
	}
}
