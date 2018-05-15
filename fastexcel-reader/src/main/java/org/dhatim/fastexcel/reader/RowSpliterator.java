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
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Optional;
import java.util.Spliterator;
import java.util.function.Consumer;
import java.util.function.Function;
import javax.xml.stream.XMLStreamException;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

class RowSpliterator implements Spliterator<Row> {

    private final SimpleXmlReader r;
    private final ReadableWorkbook workbook;

    private final HashMap<CellRangeAddress, String> sharedFormula = new HashMap<>();
    private int rowCapacity = 16;

    public RowSpliterator(ReadableWorkbook workbook, InputStream inputStream) throws XMLStreamException {
        this.workbook = workbook;
        this.r = new SimpleXmlReader(workbook.getXmlFactory(), inputStream);

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
        ArrayList<Cell> cells = new ArrayList<>(rowCapacity);
        int physicalCellCount = 0;

        while (r.goTo(() -> r.isStartElement("c") || r.isEndElement("row"))) {
            if ("row".equals(r.getLocalName())) {
                break;
            }
            String cellRef = r.getAttribute("r");
            CellAddress addr = new CellAddress(cellRef);
            String type = r.getOptionalAttribute("t").orElse("n");

            Object value;
            CellType cellType;
            String formula;
            String rawValue;
            if ("inlineStr".equals(type)) {
                Object v = null;
                String f = null;
                String rv = null;
                while (r.goTo(() -> r.isStartElement("is") || r.isEndElement("c") || r.isStartElement("f"))) {
                    if ("is".equals(r.getLocalName())) {
                        rv = r.getValueUntilEndElement("is");
                        XSSFRichTextString rtss = new XSSFRichTextString(rv);
                        v = rtss.toString();
                    } else if ("f".equals(r.getLocalName())) {
                        f = r.getValueUntilEndElement("f");
                    } else {
                        break;
                    }
                }
                value = v;
                formula = f;
                cellType = f == null ? CellType.STRING : CellType.FORMULA;
                rawValue = rv;
            } else if ("s".equals(type)) {
                r.goTo("v");
                int index = Integer.parseInt(r.getValueUntilEndElement("v"));
                CTRst ctrst = workbook.getSharedStringsTable().getEntryAt(index);
                XSSFRichTextString rtss = new XSSFRichTextString(ctrst);
                value = rtss.toString();
                cellType = CellType.STRING;
                formula = null;
                rawValue = ctrst.xmlText();
            } else {
                Function<String, ?> parser;
                CellType definedType;
                switch (type) {
                    case "b":
                        definedType = CellType.BOOLEAN;
                        parser = s -> s.charAt(0) != 0;
                        break;
                    case "e":
                        definedType = CellType.ERROR;
                        parser = Function.identity();
                        break;
                    case "n":
                        definedType = CellType.NUMBER;
                        parser = RowSpliterator::parseNumber;
                        break;
                    case "str":
                        definedType = CellType.FORMULA;
                        parser = Function.identity();
                        break;
                    default:
                        throw new IllegalStateException("Unknown cell type : " + type);
                }

                Object v = null;
                String f = null;
                String rv = null;
                while (r.goTo(() -> r.isStartElement("v") || r.isEndElement("c") || r.isStartElement("f"))) {
                    if ("v".equals(r.getLocalName())) {
                        rv = r.getValueUntilEndElement("v");
                        v = "".equals(rv) ? null : parser.apply(rv);
                    } else if ("f".equals(r.getLocalName())) {
                        String ref = r.getAttribute("ref");
                        String t = r.getAttribute("t");
                        f = r.getValueUntilEndElement("f");
                        if ("array".equals(t) && ref != null) {
                            CellRangeAddress range = CellRangeAddress.valueOf(ref);
                            sharedFormula.put(range, f);
                        }
                    } else {
                        break;
                    }
                }

                if (f == null) {
                    f = getSharedFormula(addr).orElse(null);
                }

                if (f == null && definedType == CellType.NUMBER && v == null) {
                    cellType = CellType.EMPTY;
                    formula = null;
                    value = null;
                    rawValue = rv;
                } else {
                    cellType = f == null ? definedType : CellType.FORMULA;
                    formula = f;
                    value = v;
                    rawValue = rv;
                }

            }

            if (addr.getColumn() >= cells.size()) {
                setSize(cells, addr.getColumn() + 1);
            }

            Cell cell = new Cell(workbook, cellType, value, addr, formula, rawValue);
            cells.set(addr.getColumn(), cell);
            physicalCellCount++;
        }
        rowCapacity = Math.max(rowCapacity, cells.size());
        return new Row(rowIndex, physicalCellCount, cells.toArray(new Cell[cells.size()]));
    }

    private Optional<String> getSharedFormula(CellAddress addr) {
        for (Map.Entry<CellRangeAddress, String> entry : sharedFormula.entrySet()) {
            if (entry.getKey().isInRange(addr.getRow(), addr.getColumn())) {
                return Optional.of(entry.getValue());
            }
        }
        return Optional.empty();
    }

    private static BigDecimal parseNumber(String s) {
        try {
            return new BigDecimal(s);
        } catch (NumberFormatException e) {
            throw new ExcelReaderException("Cannot parse number : " + s, e);
        }
    }

    private static void setSize(List<?> list, int newSize) {
        if (list.size() != newSize) {
            if (list.size() < newSize) {
                int toAdd = newSize - list.size();
                for (int i=0; i<toAdd; i++) {
                    list.add(null);
                }
            } else {
                for (int i=list.size() - 1; i > newSize; i--) {
                    list.remove(i);
                }
            }
        }
    }

}
