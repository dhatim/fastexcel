package org.dhatim.fastexcel.reader;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

final class FormulaResolver {

    private final Map<Integer, BaseFormulaCell> sharedFormula = new HashMap<>();
    private final Map<CellRangeAddress, String> arrayFormula = new HashMap<>();

    void registerArrayFormula(String ref, String formula) {
        arrayFormula.put(CellRangeAddress.valueOf(ref), formula);
    }

    void registerSharedFormula(Integer sharedIndex, CellAddress address, String formula, String ref) {
        sharedFormula.put(sharedIndex, new BaseFormulaCell(address, formula, CellRangeAddress.valueOf(ref)));
    }

    String resolveSharedFormula(Integer sharedIndex, CellAddress address) {
        BaseFormulaCell baseFormulaCell = sharedFormula.get(sharedIndex);
        int dRow = address.getRow() - baseFormulaCell.getBaseCelAddr().getRow();
        int dCol = address.getColumn() - baseFormulaCell.getBaseCelAddr().getColumn();
        return translateSharedFormula(dCol, dRow, baseFormulaCell.getFormula());
    }

    Optional<String> resolveArrayFormula(CellAddress address) {
        for (Map.Entry<CellRangeAddress, String> entry : arrayFormula.entrySet()) {
            if (entry.getKey().isInRange(address.getRow(), address.getColumn())) {
                return Optional.of(entry.getValue());
            }
        }
        return Optional.empty();
    }

    /**
     * @see <a href="https://github.com/qax-os/excelize/blob/master/cell.go">here</a>
     */
    private String translateSharedFormula(Integer dCol, Integer dRow, String baseFormula) {
        StringBuilder res = new StringBuilder();
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

                res.append(baseFormula.substring(start, end));
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
                    res.append(shiftCell(cellID, dCol, dRow));
                    start = end;
                }
            }
        }

        if (start < baseFormula.length()) {
            res.append(baseFormula.substring(start));
        }

        return res.toString();
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
            fCol += dCol;
        }
        if (cellID.lastIndexOf("$") > 0) {
            signRow = "$";
        } else {
            fRow += dRow;
        }
        String colName = CellAddress.convertNumToColString(fCol);
        return signCol + colName + signRow + (++fRow);
    }
}
