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

import javax.xml.stream.XMLStreamException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;

import static org.dhatim.fastexcel.reader.DefaultXMLInputFactory.factory;

class RowSpliterator implements Spliterator<Row> {

    private final SimpleXmlReader r;
    private final ReadableWorkbook workbook;

    private final HashMap<CellRangeAddress, String> sharedFormula = new HashMap<>();
    private int rowCapacity = 16;

    public RowSpliterator(ReadableWorkbook workbook, InputStream inputStream) throws XMLStreamException {
        this.workbook = workbook;
        this.r = new SimpleXmlReader(factory, inputStream);

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
        int rowIndex = r.getIntAttribute("r");
        List<Cell> cells = new ArrayList<>(rowCapacity);
        int physicalCellCount = 0;

        while (r.goTo(() -> r.isStartElement("c") || r.isEndElement("row"))) {
            if ("row".equals(r.getLocalName())) {
                break;
            }

            Cell cell = parseCell();
            CellAddress addr = cell.getAddress();
            ensureSize(cells, addr.getColumn() + 1);

            cells.set(addr.getColumn(), cell);
            physicalCellCount++;
        }
        rowCapacity = Math.max(rowCapacity, cells.size());
        return new Row(rowIndex, physicalCellCount, cells);
    }

    private Cell parseCell() throws XMLStreamException {
        String cellRef = r.getAttribute("r");
        CellAddress addr = new CellAddress(cellRef);
        String type = r.getOptionalAttribute("t").orElse("n");
        if ("inlineStr".equals(type)) {
            return parseInlineStr(addr);
        } else if ("s".equals(type)) {
            return parseString(addr);
        } else {
            return parseOther(addr, type);
        }
    }

    private Cell parseOther(CellAddress addr, String type) throws XMLStreamException {
        CellType definedType = parseType(type);
        Function<String, ?> parser = getParserForType(definedType);

        Object value = null;
        String formula = null;
        String rawValue = null;
        while (r.goTo(() -> r.isStartElement("v") || r.isEndElement("c") || r.isStartElement("f"))) {
            if ("v".equals(r.getLocalName())) {
                rawValue = r.getValueUntilEndElement("v");
                value = "".equals(rawValue) ? null : parser.apply(rawValue);
            } else if ("f".equals(r.getLocalName())) {
                String ref = r.getAttribute("ref");
                String t = r.getAttribute("t");
                formula = r.getValueUntilEndElement("f");
                if ("array".equals(t) && ref != null) {
                    CellRangeAddress range = CellRangeAddress.valueOf(ref);
                    sharedFormula.put(range, formula);
                }
            } else {
                break;
            }
        }

        if (formula == null) {
            formula = getSharedFormula(addr).orElse(null);
        }

        if (formula == null && value == null && definedType == CellType.NUMBER) {
            return new Cell(workbook, CellType.EMPTY, null, addr, null, rawValue);
        } else {
            CellType cellType = (formula != null) ? CellType.FORMULA : definedType;
            return new Cell(workbook, cellType, value, addr, formula, rawValue);
        }
    }

    private Cell parseString(CellAddress addr) throws XMLStreamException {
        r.goTo("v");
        String v = r.getValueUntilEndElement("v");
        if(v.isEmpty()) {
            return new Cell(workbook, CellType.STRING, "", addr, null, "");
        }
        int index = Integer.parseInt(v);
        String sharedStringValue = workbook.getSharedStringsTable().getItemAt(index);
        Object value = sharedStringValue;
        String formula = null;
        String rawValue = sharedStringValue;
        return new Cell(workbook, CellType.STRING, value, addr, formula, rawValue);
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

    private Optional<String> getSharedFormula(CellAddress addr) {
        for (Map.Entry<CellRangeAddress, String> entry : sharedFormula.entrySet()) {
            if (entry.getKey().isInRange(addr.getRow(), addr.getColumn())) {
                return Optional.of(entry.getValue());
            }
        }
        return Optional.empty();
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

}
