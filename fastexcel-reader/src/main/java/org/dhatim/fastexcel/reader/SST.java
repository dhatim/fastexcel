package org.dhatim.fastexcel.reader;


import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
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

  public SST(XMLInputFactory factory, InputStream in) throws XMLStreamException {
    reader = new SimpleXmlReader(factory, in);
  }

  public static SST fromInputStream(XMLInputFactory factory, InputStream in) throws XMLStreamException {
    return in == null ? EMPTY : new SST(factory, in);
  }

  public String getItemAt(int index) throws XMLStreamException {
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
