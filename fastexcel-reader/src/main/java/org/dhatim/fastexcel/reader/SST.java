package org.dhatim.fastexcel.reader;


import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLInputFactory;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

class SST {
  private static final SST EMPTY = new SST();
  private final SimpleXmlReader reader;
  private final List<String> values = new ArrayList<>();

  private SST() {
    reader = null;
  }

  SST(InputStream in, XMLInputFactory xmlInputFactory) throws XMLStreamException {
    reader = new SimpleXmlReader(xmlInputFactory, in);
  }

  static SST fromInputStream(InputStream in, XMLInputFactory xmlInputFactory) throws XMLStreamException {
    return in == null ? EMPTY : new SST(in, xmlInputFactory);
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
            StringBuilder sb = new StringBuilder();
            while (reader.goTo(() -> reader.isStartElement("t")
                    || reader.isStartElement("rPh")
                    || reader.isEndElement("si"))) {
                if (reader.isStartElement("t")) {
                    sb.append(reader.getValueUntilEndElement("t"));
                } else if (reader.isEndElement("si")) {
                    break;
                } else if (reader.isStartElement("rPh")) {
                    reader.goTo(() -> reader.isEndElement("rPh"));
                }
            }
            values.add(sb.toString());
        }
    }
}
