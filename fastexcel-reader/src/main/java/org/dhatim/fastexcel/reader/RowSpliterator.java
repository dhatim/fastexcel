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
package org.dhatim.fastexcel.reader;

import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Spliterator;
import java.util.function.Consumer;
import java.util.function.Function;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;

class RowSpliterator implements Spliterator<Row> {

    private final SimpleXmlReader r;
    private final ReadableWorkbook workbook;

    private final FormulaResolver formulaResolver = new FormulaResolver();
    private int rowCapacity = 16;
    private int trackedRowIndex = 0;

    public RowSpliterator(ReadableWorkbook workbook, InputStream inputStream) throws XMLStreamException {
        this(workbook, inputStream, DefaultXMLInputFactory.create());
    }

    public RowSpliterator(ReadableWorkbook workbook, InputStream inputStream, XMLInputFactory xmlInputFactory) throws XMLStreamException {
        this.workbook = workbook;
        this.r = new SimpleXmlReader(xmlInputFactory, inputStream);

        r.goTo("sheetData");
    }

    @Override
    public boolean tryAdvance(Consumer<? super Row> action) {
        try {
            if (hasNext()) {
                action.accept(next());
                return true;
            } else {
                return false;
            }
        } catch (XMLStreamException e) {
            throw new ExcelReaderException(e);
        }
    }

    @Override
    public Spliterator<Row> trySplit() {
        return null;
    }

    @Override
    public long estimateSize() {
        return Long.MAX_VALUE;
    }

    @Override
    public int characteristics() {
        return DISTINCT | IMMUTABLE | NONNULL | ORDERED;
    }

    private boolean hasNext() throws XMLStreamException {
        if (r.goTo(() -> r.isStartElement("row") || r.isEndElement("sheetData"))) {
            return "row".equals(r.getLocalName());
        } else {
            return false;
        }
    }


    private Row next() throws XMLStreamException {
        if (!"row".equals(r.getLocalName())) {
            throw new NoSuchElementException();
        }

        int trackedColIndex = 0;
        int rowIndex = getRowIndexWithFallback(trackedRowIndex);
        String hiddenAttribute = r.getAttribute("hidden");
        boolean isHidden = "1".equals(hiddenAttribute) || "true".equals(hiddenAttribute);

        List<Cell> cells = new ArrayList<>(rowCapacity);
        int physicalCellCount = 0;

        while (r.goTo(() -> r.isStartElement("c") || r.isEndElement("row"))) {
            if ("row".equals(r.getLocalName())) {
                break;
            }

            Cell cell = parseCell(trackedColIndex++);
            CellAddress addr = cell.getAddress();
            // we may have to adjust because we may have skipped blanks
            trackedColIndex = addr.getColumn() + 1;
            ensureSize(cells, addr.getColumn() + 1);

            cells.set(addr.getColumn(), cell);
            physicalCellCount++;
        }
        trackedRowIndex++;
        rowCapacity = Math.max(rowCapacity, cells.size());
        return new Row(rowIndex, physicalCellCount, cells, isHidden);
    }

    private int getRowIndexWithFallback(int fallbackRowIndex) {
        Integer rowIndexOrNull = r.getIntAttribute("r");
        return rowIndexOrNull != null ? rowIndexOrNull : fallbackRowIndex;
    }

    private CellAddress getCellAddressWithFallback(int trackedColIndex, int trackedRowIndex) {
        String cellRefOrNull = r.getAttribute("r");
        return cellRefOrNull != null ?
                new CellAddress(cellRefOrNull) :
                new CellAddress(trackedRowIndex, trackedColIndex);
    }

    private Cell parseCell(int trackedColIndex) throws XMLStreamException {
        CellAddress addr = getCellAddressWithFallback(trackedColIndex, trackedRowIndex);
        String type = r.getOptionalAttribute("t").orElse("n");
        String styleString = r.getAttribute("s");
        String formatId = null;
        String formatString = null;
        if (styleString != null) {
            int index = Integer.parseInt(styleString);
            formatId = workbook.getFormatId(index);
            if (formatId != null) {
                formatString = workbook.getFormatString(formatId);
            }
        }

        if ("inlineStr".equals(type)) {
            return parseInlineStr(addr);
        } else if ("s".equals(type)) {
            return parseString(addr);
        } else {
            return parseOther(addr, type, formatId, formatString);
        }
    }

    private Cell parseOther(CellAddress addr, String type, String dataFormatId, String dataFormatString)
            throws XMLStreamException {
        CellType definedType = parseType(type);
        Function<String, ?> parser = getParserForType(definedType);
        ParsedCellContent content = parseCellContent(addr, definedType, parser);

        if (content.formula == null && content.value == null && content.type == CellType.NUMBER) {
            return new Cell(workbook, CellType.EMPTY, null, addr, null, content.rawValue);
        } else {
            CellType cellType = content.formula != null ? CellType.FORMULA : content.type;
            return new Cell(workbook, cellType, content.value, addr, content.formula, content.rawValue, dataFormatId, dataFormatString);
        }
    }

    private ParsedCellContent parseCellContent(CellAddress addr, CellType definedType, Function<String, ?> parser)
            throws XMLStreamException {
        Object value = null;
        String formula = null;
        String rawValue = null;
        while (r.goTo(() -> r.isStartElement("v") || r.isEndElement("c") || r.isStartElement("f"))) {
            if ("v".equals(r.getLocalName())) {
                rawValue = r.getValueUntilEndElement("v");
                try {
                    value = "".equals(rawValue) ? null : parser.apply(rawValue);
                } catch (ExcelReaderException e) {
                    if (workbook.getReadingOptions().isCellInErrorIfParseError()) {
                        definedType = CellType.ERROR;
                    } else {
                        throw e;
                    }
                }
            } else if ("f".equals(r.getLocalName())) {
                String ref = r.getAttribute("ref");
                String t = r.getAttribute("t");
                String si = r.getAttribute("si");
                Integer siInt = si == null ? null : Integer.parseInt(si);
                formula = r.getValueUntilEndElement("f");
                if ("array".equals(t) && ref != null) {
                    formulaResolver.registerArrayFormula(ref, formula);
                }
                if ("shared".equals(t)) {
                    if (ref != null) {
                        formulaResolver.registerSharedFormula(siInt, addr, formula, ref);
                    } else {
                        formula = formulaResolver.resolveSharedFormula(siInt, addr);
                    }
                }
            } else {
                break;
            }
        }

        if (formula == null || "".equals(formula)) {
            formula = formulaResolver.resolveArrayFormula(addr).orElse(null);
        }

        return new ParsedCellContent(definedType, value, formula, rawValue);
    }

    private Cell parseString(CellAddress addr) throws XMLStreamException {
        r.goTo(() -> r.isStartElement("v") || r.isEndElement("c"));
        if (r.isEndElement("c")) {
            return empty(addr, CellType.STRING);
        }
        String v = r.getValueUntilEndElement("v");
        if (v.isEmpty()) {
            return empty(addr, CellType.STRING);
        }
        int index = Integer.parseInt(v);
        String sharedStringValue = workbook.getSharedStringsTable().getItemAt(index);
        Object value = sharedStringValue;
        String formula = null;
        String rawValue = sharedStringValue;
        return new Cell(workbook, CellType.STRING, value, addr, formula, rawValue);
    }

    private Cell empty(CellAddress addr, CellType type) {
        return new Cell(workbook, type, "", addr, null, "");
    }

    private Cell parseInlineStr(CellAddress addr) throws XMLStreamException {
        Object value = null;
        String formula = null;
        String rawValue = null;
        while (r.goTo(() -> r.isStartElement("is") || r.isEndElement("c") || r.isStartElement("f"))) {
            if ("is".equals(r.getLocalName())) {
                rawValue = r.getValueUntilEndElement("is");
                value = rawValue;
            } else if ("f".equals(r.getLocalName())) {
                formula = r.getValueUntilEndElement("f");
            } else {
                break;
            }
        }
        CellType cellType = formula == null ? CellType.STRING : CellType.FORMULA;
        return new Cell(workbook, cellType, value, addr, formula, rawValue);
    }

    private CellType parseType(String type) {
        switch (type) {
            case "b":
                return CellType.BOOLEAN;
            case "e":
                return CellType.ERROR;
            case "n":
                return CellType.NUMBER;
            case "str":
                return CellType.FORMULA;
            case "s":
            case "inlineStr":
                return CellType.STRING;
        }
        throw new IllegalStateException("Unknown cell type : " + type);
    }

    private Function<String, ?> getParserForType(CellType type) {
        switch (type) {
            case BOOLEAN:
                return RowSpliterator::parseBoolean;
            case NUMBER:
                return RowSpliterator::parseNumber;
            case FORMULA:
            case ERROR:
                return Function.identity();
        }
        throw new IllegalStateException("No parser defined for type " + type);
    }

    private static BigDecimal parseNumber(String s) {
        try {
            return new BigDecimal(s);
        } catch (NumberFormatException e) {
            throw new ExcelReaderException("Cannot parse number : " + s, e);
        }
    }

    private static Boolean parseBoolean(String s) {
        if ("0".equals(s)) {
            return Boolean.FALSE;
        } else if ("1".equals(s)) {
            return Boolean.TRUE;
        } else {
            throw new ExcelReaderException("Invalid boolean cell value: '" + s + "'. Expecting '0' or '1'.");
        }
    }

    private static void ensureSize(List<?> list, int newSize) {
        if (list.size() == newSize) {
            return;
        }
        int toAdd = newSize - list.size();
        for (int i = 0; i < toAdd; i++) {
            list.add(null);
        }
    }

    private static final class ParsedCellContent {
        private final CellType type;
        private final Object value;
        private final String formula;
        private final String rawValue;

        private ParsedCellContent(CellType type, Object value, String formula, String rawValue) {
            this.type = type;
            this.value = value;
            this.formula = formula;
            this.rawValue = rawValue;
        }
    }

}
