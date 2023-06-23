package org.dhatim.fastexcel;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

class XmlEscapeHelperTest {

	@ParameterizedTest
	@CsvSource({
			"<tag>,&lt;tag&gt;",
			"random&more,random&amp;more",
			"one'one,one&apos;one",
			"with\"signs\",with&quot;signs&quot;",
			"random&more,random&amp;more",
			"<this will be escaped \ud83d\ude01>,&lt;this will be escaped &#x1f601;&gt;",
			"nothing+!()happens,nothing+!()happens"})
	public void testEscaping(String input, String expected) {
		XmlEscapeHelper xmlEscapeHelper = new XmlEscapeHelper();
		assertEquals(expected, xmlEscapeHelper.escape(input));
	}
}
