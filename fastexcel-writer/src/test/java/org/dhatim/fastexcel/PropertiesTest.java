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
package org.dhatim.fastexcel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import java.io.ByteArrayInputStream;
import java.math.BigDecimal;
import java.time.Instant;
import java.util.function.Consumer;

import static org.assertj.core.api.Assertions.assertThat;

class PropertiesTest {
    private static final String PROPERTY_NAME = "A & \"B\" <C> 'D' &lt;";

    @Test
    void textPropertyNameIsEscaped() throws Exception {
        assertPropertyRoundTrip(properties -> properties.setTextProperty(PROPERTY_NAME, "Text & value"),
                property -> assertThat(property.getLpwstr()).isEqualTo("Text & value"));
    }

    @Test
    void datePropertyNameIsEscaped() throws Exception {
        Instant value = Instant.parse("2026-01-02T03:04:05Z");
        assertPropertyRoundTrip(properties -> properties.setDateProperty(PROPERTY_NAME, value),
                property -> assertThat(property.getFiletime().toInstant()).isEqualTo(value));
    }

    @Test
    void numberPropertyNameIsEscaped() throws Exception {
        assertPropertyRoundTrip(properties -> properties.setNumberProperty(PROPERTY_NAME, new BigDecimal("123.5")),
                property -> assertThat(property.getR8()).isEqualTo(123.5));
    }

    @Test
    void boolPropertyNameIsEscaped() throws Exception {
        assertPropertyRoundTrip(properties -> properties.setBoolProperty(PROPERTY_NAME, true),
                property -> assertThat(property.getBool()).isTrue());
    }

    private void assertPropertyRoundTrip(Consumer<Properties> configure, Consumer<CTProperty> assertValue)
            throws Exception {
        byte[] bytes = CorrectnessTest.writeWorkbook(workbook -> {
            workbook.newWorksheet("Sheet1");
            configure.accept(workbook.properties());
        });
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            CTProperty property = workbook.getProperties().getCustomProperties().getProperty(PROPERTY_NAME);
            assertThat(property).isNotNull();
            assertThat(property.getName()).isEqualTo(PROPERTY_NAME);
            assertValue.accept(property);
        }
    }
}
