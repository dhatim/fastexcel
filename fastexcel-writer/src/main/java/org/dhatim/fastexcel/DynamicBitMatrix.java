package org.dhatim.fastexcel;

import java.util.concurrent.CopyOnWriteArrayList;

/**
 * This is a dynamically expanding matrix structure that saves space and has good performance
 *
 * @author meiMingle
 */
public class DynamicBitMatrix {
    static final int UNIT_WEITH = 1 << 6, UNIT_HIGHT = 1 << 10;

    CopyOnWriteArrayList<CopyOnWriteArrayList<BitMatrix>> bitMatrixData = new CopyOnWriteArrayList<>();

    void setRegion(int top, int left, int bottom, int right) {
        if (right >= Worksheet.MAX_COLS || bottom >= Worksheet.MAX_ROWS) {
            throw new IllegalArgumentException();
        }
        int rightBitMatrixColIndex = right / UNIT_WEITH;
        int leftBitMatrixColIndex = left / UNIT_WEITH;
        int topBitMatrixRowIndex = top / UNIT_HIGHT;
        int bottomBitMatrixRowIndex = bottom / UNIT_HIGHT;
        if (rightBitMatrixColIndex >= bitMatrixData.size()) {
            for (int i = bitMatrixData.size() - 1; i < rightBitMatrixColIndex; i++) {
                bitMatrixData.add(null);
            }
        }
        for (int i = leftBitMatrixColIndex; i <= rightBitMatrixColIndex; i++) {
            if (bitMatrixData.get(i) == null || bitMatrixData.get(i).isEmpty()) {
                bitMatrixData.set(i, new CopyOnWriteArrayList<>());
            }
            CopyOnWriteArrayList<BitMatrix> colBitMatrices = bitMatrixData.get(i);
            if (bottomBitMatrixRowIndex >= colBitMatrices.size()) {
                for (int j = colBitMatrices.size() - 1; j < bottomBitMatrixRowIndex; j++) {
                    colBitMatrices.add(null);
                }
            }
            for (int j = topBitMatrixRowIndex; j <= bottomBitMatrixRowIndex; j++) {
                if (colBitMatrices.get(j) == null || colBitMatrices.isEmpty()) {
                    colBitMatrices.set(j, new BitMatrix(UNIT_WEITH, UNIT_HIGHT));
                }
                BitMatrix bitMatrix = colBitMatrices.get(j);

                int l = Math.max(i * UNIT_WEITH, left) - i * UNIT_WEITH;
                int t = Math.max(j * UNIT_HIGHT, top) - j * UNIT_HIGHT;
                int r = Math.min((i + 1) * UNIT_WEITH - 1, right) - i * UNIT_WEITH;
                int b = Math.min((j + 1) * UNIT_HIGHT - 1, bottom) - j * UNIT_HIGHT;

                bitMatrix.setRegion(l, t, r - l + 1, b - t + 1);
            }


        }

    }

    boolean get(int row, int col) {
        int bitMatrixColIndex = col / UNIT_WEITH;
        int bitMatrixRowIndex = row / UNIT_HIGHT;
        return !isInNullArea(bitMatrixRowIndex, bitMatrixColIndex) && bitMatrixData.get(bitMatrixColIndex).get(bitMatrixRowIndex).get(col-bitMatrixColIndex * UNIT_WEITH, row - bitMatrixRowIndex * UNIT_HIGHT);
    }

    private boolean isInNullArea(int bitMatrixRowIndex, int bitMatrixColIndex) {
        if (bitMatrixColIndex >= bitMatrixData.size()) {
            return true;
        }
        CopyOnWriteArrayList<BitMatrix> colBitMatrices = bitMatrixData.get(bitMatrixColIndex);
        if (colBitMatrices == null || colBitMatrices.isEmpty()) {
            return true;
        }
        if (bitMatrixRowIndex >= colBitMatrices.size()) {
            return true;
        }
        return colBitMatrices.get(bitMatrixRowIndex) == null;
    }


    @Override
    public String toString() {
        return buildToString("1", "0", " ", "\n");
    }

    public String buildToString(String setString, String unsetString, String fillNullString ,String lineSeparator) {
        StringBuilder builder = new StringBuilder();
        int maxBitMatrixCol = bitMatrixData.size();
        int maxBitMatrixRow = bitMatrixData.stream().mapToInt(a -> a == null ? 0 : a.size()).max().orElse(0);
        for (int i = 0; i < maxBitMatrixRow; i++) {
            for (int h = 0; h < UNIT_HIGHT; h++) {
                for (int j = 0; j < maxBitMatrixCol; j++) {
                    boolean inNullArea = isInNullArea(i, j);
                        for (int k = 0; k < UNIT_WEITH; k++) {
                            builder.append(inNullArea ? fillNullString : bitMatrixData.get(j).get(i).get(k, h) ? setString : unsetString);
                            builder.append(j == maxBitMatrixCol - 1 && k == UNIT_WEITH - 1 ? lineSeparator : ',');
                        }
                }
            }

        }
        return builder.toString();
    }

}
