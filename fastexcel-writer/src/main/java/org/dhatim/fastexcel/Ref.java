package org.dhatim.fastexcel;

public interface Ref {
    default   String colToString(int col){
        StringBuilder sb = new StringBuilder();
        while (col >= 0) {
            sb.append((char) ('A' + (col % 26)));
            col = (col / 26) - 1;
        }
        return sb.reverse().toString();
    }
}
