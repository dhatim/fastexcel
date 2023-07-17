package org.dhatim.fastexcel.reader;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

public class Resources {
    static InputStream open(String name) {
        InputStream result = Resources.class.getResourceAsStream(name);
        if (result == null) {
            throw new IllegalStateException("Cannot read resource " + name);
        }
        return result;
    }

    static File file(String name) {
        try {
            File tempFile = File.createTempFile(name, ".tmp");
            tempFile.deleteOnExit();
            try (InputStream in = open(name)) {
                Files.copy(in, tempFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            }
            return tempFile;
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
}
