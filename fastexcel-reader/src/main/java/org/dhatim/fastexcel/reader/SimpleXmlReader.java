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

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.util.Optional;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;

class SimpleXmlReader implements Closeable {

    private final InputStream inputStream;
    private final XMLStreamReader reader;

    public SimpleXmlReader(XMLInputFactory factory, InputStream inputStream) throws XMLStreamException {
        this.inputStream = inputStream;
        reader = factory.createXMLStreamReader(inputStream);
    }

    @Override
    public void close() throws IOException {
        inputStream.close();
    }

    public boolean goTo(BooleanSupplier predicate) throws XMLStreamException {
        while (reader.hasNext()) {
            reader.next();
            if (predicate.getAsBoolean()) {
                return true;
            }
        }
        return false;
    }

    public String getLocalName() {
        return reader.getLocalName();
    }

    public boolean isStartElement(String elementName) {
        return reader.isStartElement() && elementName.equals(reader.getLocalName());
    }

    public boolean isEndElement(String elementName) {
        return reader.isEndElement() && elementName.equals(reader.getLocalName());
    }

    public boolean goTo(String elementName) throws XMLStreamException {
        return goTo(() -> isStartElement(elementName));
    }

    public String getAttribute(String name) {
        return reader.getAttributeValue(null, name);
    }

    public String getAttributeRequired(String name) throws XMLStreamException {
        String value = getAttribute(name);
        if(value == null) {
            throw new XMLStreamException("missing required attribute "+name);
        }
        return value;
    }

    public String getAttribute(String namespace, String name) {
        return reader.getAttributeValue(namespace, name);
    }

    public Optional<String> getOptionalAttribute(String name) {
        return Optional.ofNullable(reader.getAttributeValue(null, name));
    }

    public Integer getIntAttribute(String name) {
        String value = reader.getAttributeValue(null, name);
        return value == null ? null : Integer.valueOf(value);
    }

    public void forEach(String startChildElement, String untilEndElement, Consumer<SimpleXmlReader> consumer) throws XMLStreamException {
        while (goTo(() -> isStartElement(startChildElement) || isEndElement(untilEndElement))) {
            if (untilEndElement.equals(getLocalName())) {
                break;
            }
            consumer.accept(this);
        }
    }

    public String getValueUntilEndElement(String elementName) throws XMLStreamException {
        return getValueUntilEndElement(elementName, "");
    }

    public String getValueUntilEndElement(String elementName, String skipping) throws XMLStreamException {
        StringBuilder sb = new StringBuilder();
        int childElement = 1;
        while (reader.hasNext()) {
            int type = reader.next();
            if (type == XMLStreamReader.CDATA || type == XMLStreamReader.CHARACTERS || type == XMLStreamReader.SPACE) {
                sb.append(reader.getText());
            } else if (type == XMLStreamReader.START_ELEMENT) {
                if(skipping.equals(reader.getLocalName())) {
                    getValueUntilEndElement(reader.getLocalName());
                }else {
                    childElement++;
                }
            } else if (type == XMLStreamReader.END_ELEMENT) {
                childElement--;
                if (elementName.equals(reader.getLocalName()) && childElement == 0) {
                    break;
                }
            }
        }
        return sb.toString();
    }
}
