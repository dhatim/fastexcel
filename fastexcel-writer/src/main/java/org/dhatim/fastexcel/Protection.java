package org.dhatim.fastexcel;

import java.io.IOException;
import java.util.Map;
import java.util.Objects;

/**
 * Represents the &lt;protection&gt; xml-tag.
 */
public class Protection {

    private final Map<ProtectionOption, Boolean> options;

    public Protection(Map<ProtectionOption, Boolean> options) {
        if (options == null) {
            throw new NullPointerException("Options should not be null");
        }
        this.options = options;
    }

    @Override
    public int hashCode() {
        return options.hashCode();
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Protection that = (Protection) o;
        return Objects.equals(options, that.options);
    }

    /**
     * Write this protection as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    void write(Writer w) throws IOException {
        w.append("<protection ");
        for (Map.Entry<ProtectionOption, Boolean> option : options.entrySet()) {
            w.append(option.getKey().getName()).append("=\"").append(option.getValue().toString()).append("\" ");
        }
        w.append("/>");
    }

}
