package org.dhatim.fastexcel.reader;

import org.junit.jupiter.api.Test;

import java.io.InputStream;
import java.util.Objects;
import java.util.stream.Stream;

import static java.util.stream.Collectors.joining;
import static org.assertj.core.api.Assertions.assertThat;

public class ParallelStreamTest {

  @Test
  void parallelStreams() throws Exception {
    try (InputStream in = Resources.open("/xlsx/calendar_stress_test.xlsx");
         ReadableWorkbook wb = new ReadableWorkbook(in)) {
      String sequential = wb.getFirstSheet().openStream()
          .map(row -> collectToString(row.stream()))
          .collect(joining("\n"));
      String parallel = wb.getFirstSheet().openStream().parallel()
          .map(row -> collectToString(row.stream().parallel()))
          .collect(joining("\n"));
      assertThat(sequential).isEqualTo(parallel);
    }
  }

  private String collectToString(Stream<Cell> stream) {
    return stream
        .filter(Objects::nonNull)
        .map(Cell::getValue)
        .map(String::valueOf)
        .collect(joining(";"));
  }
}
