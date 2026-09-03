package org.dhatim.fastexcel.reader;

import java.util.Collections;
import java.util.List;
import java.util.Map;

final class FormatCatalog {

    static final FormatCatalog EMPTY = new FormatCatalog(Collections.emptyList(), Collections.emptyMap());

    private final List<String> formatIdsByStyleIndex;
    private final Map<String, String> formatsById;

    FormatCatalog(List<String> formatIdsByStyleIndex, Map<String, String> formatsById) {
        this.formatIdsByStyleIndex = Collections.unmodifiableList(formatIdsByStyleIndex);
        this.formatsById = Collections.unmodifiableMap(formatsById);
    }

    String getFormatId(int styleIndex) {
        return styleIndex < formatIdsByStyleIndex.size() ? formatIdsByStyleIndex.get(styleIndex) : null;
    }

    String getFormatString(String formatId) {
        return formatsById.get(formatId);
    }

    List<String> getFormatIdsByStyleIndex() {
        return formatIdsByStyleIndex;
    }

    Map<String, String> getFormatsById() {
        return formatsById;
    }
}
