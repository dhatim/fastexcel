package org.dhatim.fastexcel.reader;

import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Sheet {

    private final ReadableWorkbook workbook;
    private final int index;
    private final String id;
    private final String name;

    public Sheet(ReadableWorkbook workbook, int index, String id, String name) {
        this.workbook = workbook;
        this.index = index;
        this.id = id;
        this.name = name;
    }

    public int getIndex() {
        return index;
    }

    public String getId() {
        return id;
    }

    public String getName() {
        return name;
    }

    public Stream<Row> openStream() throws IOException {
        return workbook.openStream(this);
    }

    public List<Row> read() throws IOException {
        try (Stream<Row> stream = openStream()) {
            return stream.collect(Collectors.toList());
        }
    }

}
