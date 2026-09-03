package org.dhatim.fastexcel;

import java.util.concurrent.CopyOnWriteArrayList;

/**
 * This is a dynamically expanding matrix structure that saves space and has good performance
 *
 * @author meiMingle
 */
public class DynamicBitMatrix {
    static final int UNIT_WEITH = 1 << 6, UNIT_HIGHT = 1 << 10;

    final int MAX_WIDTH ,MAX_HIGHT;

    private final CopyOnWriteArrayList<CopyOnWriteArrayList<BitMatrix>> bitMatrixData = new CopyOnWriteArrayList<>();

    public DynamicBitMatrix(int maxWidth, int maxHight) {
        MAX_WIDTH = maxWidth;
        MAX_HIGHT = maxHight;
    }

    void setRegion(int top, int left, int bottom, int right) {
        if (right >= MAX_WIDTH ) {
            throw new IllegalArgumentException(String.format("Right boundary value exceeds maximum allowed.MAX_WIDTH = %d,right = %d",MAX_WIDTH,right));
        }
        if (bottom >= MAX_HIGHT ) {
            throw new IllegalArgumentException(String.format("Bottom boundary value exceeds maximum allowed.MAX_HIGHT = %d,bottom = %d",MAX_HIGHT,bottom));
        }
        int rightBitMatrixColIndex = right / UNIT_WEITH;
        int leftBitMatrixColIndex = left / UNIT_WEITH;
        int topBitMatrixRowIndex = top / UNIT_HIGHT;
        int bottomBitMatrixRowIndex = bottom / UNIT_HIGHT;
        for (int i = leftBitMatrixColIndex; i <= rightBitMatrixColIndex; i++) {
            CopyOnWriteArrayList<BitMatrix> colBitMatrices = ensureColumn(i);
            for (int j = topBitMatrixRowIndex; j <= bottomBitMatrixRowIndex; j++) {
                BitMatrix bitMatrix = ensureChunk(colBitMatrices, j);
                MatrixRegion region = MatrixRegion.intersection(i, j, top, left, bottom, right);
                bitMatrix.setRegion(region.left, region.top, region.width(), region.height());
            }


        }

    }

    private CopyOnWriteArrayList<BitMatrix> ensureColumn(int columnIndex) {
        ensureSize(bitMatrixData, columnIndex);
        CopyOnWriteArrayList<BitMatrix> colBitMatrices = bitMatrixData.get(columnIndex);
        if (colBitMatrices == null || colBitMatrices.isEmpty()) {
            colBitMatrices = new CopyOnWriteArrayList<>();
            bitMatrixData.set(columnIndex, colBitMatrices);
        }
        return colBitMatrices;
    }

    private BitMatrix ensureChunk(CopyOnWriteArrayList<BitMatrix> colBitMatrices, int rowIndex) {
        ensureSize(colBitMatrices, rowIndex);
        BitMatrix bitMatrix = colBitMatrices.get(rowIndex);
        if (bitMatrix == null) {
            bitMatrix = new BitMatrix(UNIT_WEITH, UNIT_HIGHT);
            colBitMatrices.set(rowIndex, bitMatrix);
        }
        return bitMatrix;
    }

    private static <T> void ensureSize(CopyOnWriteArrayList<T> list, int index) {
        for (int i = list.size() - 1; i < index; i++) {
            list.add(null);
        }
    }

    boolean isConflict(int top, int left, int bottom, int right) {
        if (get(top, left) || get(top, right) || get(bottom, left) || get(bottom, right)) {
            return true;
        }
        for (int c = left; c <= right; c++) {
            for (int r = top; r <= bottom; r++) {
                if ((c == left && (r == top || r == bottom)) || (c == right && (r == top || r == bottom))) {
                    continue;
                }
                if (get(r, c)) {
                    return true;
                }
            }
        }
        return false;
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
                    appendChunkRow(builder, i, h, j, maxBitMatrixCol, setString, unsetString, fillNullString, lineSeparator);
                }
            }

        }
        return builder.toString();
    }

    private void appendChunkRow(StringBuilder builder, int chunkRow, int localRow, int chunkCol, int maxBitMatrixCol,
                                String setString, String unsetString, String fillNullString, String lineSeparator) {
        boolean inNullArea = isInNullArea(chunkRow, chunkCol);
        for (int k = 0; k < UNIT_WEITH; k++) {
            builder.append(inNullArea ? fillNullString : bitMatrixData.get(chunkCol).get(chunkRow).get(k, localRow) ? setString : unsetString);
            builder.append(chunkCol == maxBitMatrixCol - 1 && k == UNIT_WEITH - 1 ? lineSeparator : ',');
        }
    }

    private static final class MatrixRegion {
        private final int left;
        private final int top;
        private final int right;
        private final int bottom;

        private MatrixRegion(int left, int top, int right, int bottom) {
            this.left = left;
            this.top = top;
            this.right = right;
            this.bottom = bottom;
        }

        private static MatrixRegion intersection(int chunkCol, int chunkRow, int top, int left, int bottom, int right) {
            int chunkLeft = chunkCol * UNIT_WEITH;
            int chunkTop = chunkRow * UNIT_HIGHT;
            return new MatrixRegion(
                    Math.max(chunkLeft, left) - chunkLeft,
                    Math.max(chunkTop, top) - chunkTop,
                    Math.min((chunkCol + 1) * UNIT_WEITH - 1, right) - chunkLeft,
                    Math.min((chunkRow + 1) * UNIT_HIGHT - 1, bottom) - chunkTop);
        }

        private int width() {
            return right - left + 1;
        }

        private int height() {
            return bottom - top + 1;
        }
    }

}
