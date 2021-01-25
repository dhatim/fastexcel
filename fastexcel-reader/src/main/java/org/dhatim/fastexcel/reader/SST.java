package org.dhatim.fastexcel.reader;


import javax.xml.stream.XMLStreamException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.dhatim.fastexcel.reader.DefaultXMLInputFactory.factory;

class SST {
  private static final SST EMPTY = new SST();
  private final SimpleXmlReader reader;
  private final List<String> values = new ArrayList<>();

  private SST() {
    reader = null;
  }

  SST(InputStream in) throws XMLStreamException {
    reader = new SimpleXmlReader(factory, in);
  }

  static SST fromInputStream(InputStream in) throws XMLStreamException {
    return in == null ? EMPTY : new SST(in);
  }

  String getItemAt(int index) throws XMLStreamException {
    if (reader == null) {
      return null;
    }
    readUpTo(index);
    return values.get(index);
  }

  private void readUpTo(int index) throws XMLStreamException {
    while (index >= values.size()) {
      reader.goTo("si");
      values.add(reader.getValueUntilEndElement("si", "rPh"));
    }
  }
}
