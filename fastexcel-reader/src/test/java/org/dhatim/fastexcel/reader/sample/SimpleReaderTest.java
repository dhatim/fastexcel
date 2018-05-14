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
package org.dhatim.fastexcel.reader.sample;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.stream.Stream;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.junit.jupiter.api.Test;

public class SimpleReaderTest {

    private static final Object[][] VALUES = {
            {1, "Lorem", date(2018, 1, 1)},
            {2, "ipsum", date(2018, 1, 2)},
            {3, "dolor", date(2018, 1, 3)},
            {4, "sit", date(2018, 1, 4)},
            {5, "amet", date(2018, 1, 5)},
            {6, "consectetur", date(2018, 1, 6)},
            {7, "adipiscing", date(2018, 1, 7)},
            {8, "elit", date(2018, 1, 8)},
            {9, "Ut", date(2018, 1, 9)},
            {10, "nec", date(2018, 1, 10)},
    };

    @Test
    public void test() throws IOException {
        try (InputStream is = openResource("/xlsx/simple.xlsx"); ReadableWorkbook wb = new ReadableWorkbook(is)) {
            Sheet sheet = wb.getFirstSheet();
            try (Stream<Row> rows = sheet.openStream()) {
                rows.forEach(r -> {
                    BigDecimal num = r.getCellAsNumber(0).orElse(null);
                    String str = r.getCellAsString(1).orElse(null);
                    LocalDateTime date = r.getCellAsDate(2).orElse(null);

                    Object[] values = VALUES[r.getRowNum() - 1];
                    assertThat(num).isEqualTo(BigDecimal.valueOf((Integer) values[0]));
                    assertThat(str).isEqualTo((String) values[1]);
                    assertThat(date).isEqualTo((LocalDateTime) values[2]);
                });
            }
        }
    }

    private static InputStream openResource(String name) {
        InputStream result = SimpleReaderTest.class.getResourceAsStream(name);
        if (result == null) {
            throw new IllegalStateException("Cannot read resource " + name);
        }
        return result;
    }

    private static LocalDateTime date(int year, int month, int day) {
        return LocalDateTime.of(year, month, day, 0, 0);
    }


}
