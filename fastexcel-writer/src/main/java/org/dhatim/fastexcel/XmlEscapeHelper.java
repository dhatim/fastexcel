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
	public static String escape(final String text) {
		if (text == null || text.isEmpty()) {
			return text;
		}
		StringBuilder sb = new StringBuilder(text.length());
		appendEscaped(sb, text);
		return sb.toString();
	}

	/**
	 * Append XML-escaped text to a StringBuilder without allocating per-character Strings.
	 * Invalid characters in XML 1.0 are ignored.
	 *
	 * @param sb target to append to
	 * @param text text to be escaped
	 */
	public static void appendEscaped(StringBuilder sb, final String text) {
		if (text == null) {
			return;
		}
		int offset = 0;
		final int length = text.length();
		while (offset < length) {
			int codePoint = text.codePointAt(offset);
			appendEscapedCodePoint(sb, codePoint);
			offset += Character.charCount(codePoint);
		}
	}

	/**
	 * Append a single code point with XML escaping to the given StringBuilder.
	 * Invalid characters in XML 1.0 are skipped (nothing appended).
	 *
	 * @param sb target to append to
	 * @param c Character code point.
	 */
	private static void appendEscapedCodePoint(StringBuilder sb, int c) {
		if (!(c == 0x9 || c == 0xa || c == 0xD
				|| (c >= 0x20 && c <= 0xd7ff)
				|| (c >= 0xe000 && c <= 0xfffd)
				|| (c >= 0x10000 && c <= 0x10ffff))) {
			return;
		}
		switch (c) {
			case '<':
				sb.append("&lt;");
				break;
			case '>':
				sb.append("&gt;");
				break;
			case '&':
				sb.append("&amp;");
				break;
			case '\'':
				sb.append("&apos;");
				break;
			case '"':
				sb.append("&quot;");
				break;
			default:
				if (c > 0x7e || c < 0x20) {
					sb.append("&#x").append(Integer.toHexString(c)).append(";");
				} else {
					sb.appendCodePoint(c);
				}
				break;
		}
	}
}
