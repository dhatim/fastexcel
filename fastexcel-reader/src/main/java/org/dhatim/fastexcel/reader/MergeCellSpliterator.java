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

import static org.dhatim.fastexcel.reader.DefaultXMLInputFactory.factory;

import java.io.InputStream;
import java.util.NoSuchElementException;
import java.util.Spliterator;
import java.util.function.Consumer;
import javax.xml.stream.XMLStreamException;

class MergeCellSpliterator implements Spliterator<CellRangeAddress> {

    private final SimpleXmlReader r;

    public MergeCellSpliterator(InputStream inputStream) throws XMLStreamException {
        this.r = new SimpleXmlReader(factory, inputStream);
    }

    @Override
    public boolean tryAdvance(Consumer<? super CellRangeAddress> action) {
        try {
            if (hasNext()) {
                action.accept(next());
                return true;
            } else {
                return false;
            }
        } catch (XMLStreamException e) {
            throw new ExcelReaderException(e);
        }
    }

    @Override
    public Spliterator<CellRangeAddress> trySplit() {
        return null;
    }

    @Override
    public long estimateSize() {
        return Long.MAX_VALUE;
    }

    @Override
    public int characteristics() {
        return DISTINCT | IMMUTABLE | NONNULL | ORDERED;
    }

    private boolean hasNext() throws XMLStreamException {
        if (r.goTo(() -> r.isStartElement("mergeCell") || r.isEndElement("mergeCells"))) {
            return "mergeCell".equals(r.getLocalName());
        } else {
            return false;
        }
    }


    private CellRangeAddress next() {
        if (!"mergeCell".equals(r.getLocalName())) {
            throw new NoSuchElementException();
        }

        String ref = r.getAttribute("ref");
        return CellRangeAddress.valueOf(ref);
    }
}
