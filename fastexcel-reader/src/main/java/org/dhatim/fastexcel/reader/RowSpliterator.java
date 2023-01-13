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

    private final HashMap<Integer, BaseFormulaCell> sharedFormula = new HashMap<>();
    private final HashMap<CellRangeAddress, String> arrayFormula = new HashMap<>();
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
        String styleString = r.getAttribute("s");
        String formatId = null;
        String formatString = null;
        if (styleString != null) {
            int index = Integer.parseInt(styleString);
            if (index < workbook.getFormats().size()) {
                formatId = workbook.getFormats().get(index);
                formatString = workbook.getNumFmtIdToFormat().get(formatId);
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
                    CellRangeAddress range = CellRangeAddress.valueOf(ref);
                    arrayFormula.put(range, formula);
                }
                if ("shared".equals(t)) {
                    if (ref != null) {
                        CellRangeAddress range = CellRangeAddress.valueOf(ref);
                        sharedFormula.put(siInt, new BaseFormulaCell(addr, formula, range));
                    } else {
                        formula = parseSharedFormula(siInt, addr);
                    }
                }
            } else {
                break;
            }
        }

        if (formula == null || "".equals(formula)) {
            formula = getArrayFormula(addr).orElse(null);
        }

        if (formula == null && value == null && definedType == CellType.NUMBER) {
            return new Cell(workbook, CellType.EMPTY, null, addr, null, rawValue);
        } else {
            CellType cellType = (formula != null) ? CellType.FORMULA : definedType;
            return new Cell(workbook, cellType, value, addr, formula, rawValue, dataFormatId, dataFormatString);
        }
    }

    private String parseSharedFormula(Integer si, CellAddress addr) {
        BaseFormulaCell baseFormulaCell = sharedFormula.get(si);
        int dRow = addr.getRow() - baseFormulaCell.getBaseCelAddr().getRow();
        int dCol = addr.getColumn() - baseFormulaCell.getBaseCelAddr().getColumn();
        String baseFormula = baseFormulaCell.getFormula();
        return parseSharedFormula(dCol, dRow, baseFormula);
    }

    /**
     * @see <a href="https://github.com/qax-os/excelize/blob/master/cell.go">here</a>
     */
    private String parseSharedFormula(Integer dCol, Integer dRow, String baseFormula) {
        String res = "";
        int start = 0;
        boolean stringLiteral = false;
        for (int end = 0; end < baseFormula.length(); end++) {
            char c = baseFormula.charAt(end);
            if ('"' == c) {
                stringLiteral = !stringLiteral;
            }
            if (stringLiteral) {
                continue;// Skip characters in quotes
            }
            if (c >= 'A' && c <= 'Z' || c == '$') {

                res += baseFormula.substring(start, end);
                start = end;
                end++;
                boolean foundNum = false;
                for (; end < baseFormula.length(); end++) {
                    char idc = baseFormula.charAt(end);
                    if (idc >= '0' && idc <= '9' || idc == '$') {
                        foundNum = true;
                    } else if (idc >= 'A' && idc <= 'Z') {
                        if (foundNum) {
                            break;
                        }
                    } else {
                        break;
                    }
                }
                if (foundNum) {
                    String cellID = baseFormula.substring(start, end);
                    res += shiftCell(cellID, dCol, dRow);
                    start = end;
                }
            }
        }

        if (start < baseFormula.length()) {
            res += baseFormula.substring(start);
        }

        return res;
    }

    /**
     * @see <a href="https://github.com/qax-os/excelize/blob/master/cell.go">here</a>
     */
    private String shiftCell(String cellID, Integer dCol, Integer dRow) {
        CellAddress cellAddress = new CellAddress(cellID);
        int fCol = cellAddress.getColumn();
        int fRow = cellAddress.getRow();

        String signCol = "", signRow = "";
        if (cellID.indexOf("$") == 0) {
            signCol = "$";
        } else {
            // Shift column
            fCol += dCol;
        }
        if (cellID.lastIndexOf("$") > 0) {
            signRow = "$";
        } else {
            // Shift row
            fRow += dRow;
        }
        String colName = CellAddress.convertNumToColString(fCol);
        return signCol + colName + signRow + (++fRow);
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

    private Optional<String> getArrayFormula(CellAddress addr) {
        for (Map.Entry<CellRangeAddress, String> entry : arrayFormula.entrySet()) {
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
