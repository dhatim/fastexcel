package org.dhatim.fastexcel.reader;

import java.io.InputStream;

public class Resources {
    static InputStream open(String name) {
        InputStream result = Resources.class.getResourceAsStream(name);
        if (result == null) {
            throw new IllegalStateException("Cannot read resource " + name);
        }
        return result;
    }
}
