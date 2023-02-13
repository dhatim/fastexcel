package org.dhatim.fastexcel;

import java.util.concurrent.CopyOnWriteArrayList;

/**
 * @author meiMingle
 */
public class DynamicByteArray {

    static final int UNIT_LENGTH = 1 << 6;

    private final CopyOnWriteArrayList<byte[]> byteArrayData = new CopyOnWriteArrayList<>();

    final int MAX_LENGTH;

    public DynamicByteArray(int maxLength) {
        MAX_LENGTH = maxLength;
    }

    void set(int index, byte value) {
        if (index >= MAX_LENGTH) {
            throw new IllegalArgumentException(String.format("Index value exceeds the maximum allowed length value.MAX_LENGTH = %d,index = %d", MAX_LENGTH, index));
        }
        int arrayAreaIndex = index / UNIT_LENGTH;
        byte[] bytes = getBytesWithOutNull(arrayAreaIndex);
        bytes[index - arrayAreaIndex * UNIT_LENGTH] = value;
    }

    void increase(int index) {
        if (index >= MAX_LENGTH) {
            throw new IllegalArgumentException(String.format("Index value exceeds the maximum allowed length value.MAX_LENGTH = %d,index = %d", MAX_LENGTH, index));
        }
        int arrayAreaIndex = index / UNIT_LENGTH;
        byte[] bytes = getBytesWithOutNull(arrayAreaIndex);
        if (++bytes[index - arrayAreaIndex * UNIT_LENGTH]>7) {
            throw new IllegalStateException(String.format("Cannot proceed with grouping because the grouping level at this index position has reached the maximum 7.index = %s",index));
        }
    }

    private byte[] getBytesWithOutNull(int arrayAreaIndex) {
        if (arrayAreaIndex >= byteArrayData.size()) {
            for (int i = byteArrayData.size() - 1; i < arrayAreaIndex; i++) {
                byteArrayData.add(null);
            }
        }
        if (byteArrayData.get(arrayAreaIndex) == null) {
            byteArrayData.set(arrayAreaIndex, new byte[UNIT_LENGTH]);
        }
        return byteArrayData.get(arrayAreaIndex);
    }


    byte get(int index) {
        int arrayAreaIndex = index / UNIT_LENGTH;
        if (arrayAreaIndex > byteArrayData.size()-1){
            return 0;
        }
        byte[] bytes = byteArrayData.get(arrayAreaIndex);
        if (bytes == null) {
            return 0;
        }
        return bytes[index - arrayAreaIndex * UNIT_LENGTH];
    }

    int getMaxNoZeroIndex() {
        for (int areaIndex = byteArrayData.size(); areaIndex > 0; areaIndex--) {
            byte[] bytes = byteArrayData.get(areaIndex-1);
            if (bytes == null) {
                continue;
            }
            for (int i = bytes.length - 1; i >= 0; i--) {
                if (bytes[i] != 0) {
                    return i + (areaIndex-1) * UNIT_LENGTH;
                }
            }
        }
        return -1;
    }


    @Override
    public String toString() {
        return buildToString(" ");
    }

    public String buildToString(String fillNullString) {
        StringBuilder builder = new StringBuilder();
        for (byte[] bytes : byteArrayData) {
            if (bytes == null) {
                builder.append(repeatString(fillNullString + ',', UNIT_LENGTH));
            } else {
                for (byte aByte : bytes) {
                    builder.append(aByte).append(',');
                }
            }
        }
        return builder.toString();
    }

    private String repeatString(String seed, int n) {
        final int seedLen = seed.length();
        final char[] srcArr = seed.toCharArray();
        char[] dstArr = new char[n * seedLen];
        for (int i = 0; i < n; i++) {
            for (int j = 0; j < seedLen; j++) {
                dstArr[i * seedLen + j] = srcArr[j];
            }
        }
        return String.valueOf(dstArr);
    }

}
