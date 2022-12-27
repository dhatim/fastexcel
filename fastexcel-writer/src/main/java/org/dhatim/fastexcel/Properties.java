package org.dhatim.fastexcel;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class Properties {

    //*****  core properties  *****
    private String title;
    private String subject;
    // alias:Tags
    private String keywords;
    // alias:Comments
    private String description;
    private String category;
    //*****  app properties  *****
    private String manager;
    private String company;
    private String hyperlinkBase;

    //***** custom properties *****
    private Set<CustomProty> customProperties = Collections.synchronizedSet(new LinkedHashSet<>());

    String getTitle() {
        return title;
    }

    public Properties setTitle(String title) {
        if (title != null && title.length() > 65536) {
            throw new IllegalStateException("The length of title must be less than or equal to 65536: " + title.length());
        }
        this.title = title;
        return this;
    }

    String getSubject() {
        return subject;
    }

    public Properties setSubject(String subject) {
        if (subject != null && subject.length() > 65536) {
            throw new IllegalStateException("The length of subject must be less than or equal to 65536: " + subject.length());
        }
        this.subject = subject;
        return this;
    }

    String getKeywords() {
        return keywords;
    }

    public Properties setKeywords(String keywords) {
        if (keywords != null && keywords.length() > 65536) {
            throw new IllegalStateException("The length of keywords must be less than or equal to 65536: " + keywords.length());
        }
        this.keywords = keywords;
        return this;
    }

    String getDescription() {
        return description;
    }

    public Properties setDescription(String description) {
        if (description != null && description.length() > 65536) {
            throw new IllegalStateException("The length of description must be less than or equal to 65536: " + description.length());
        }
        this.description = description;
        return this;
    }

    String getCategory() {
        return category;
    }

    public Properties setCategory(String category) {
        if (category != null && category.length() > 65536) {
            throw new IllegalStateException("The length of category must be less than or equal to 65536: " + category.length());
        }
        this.category = category;
        return this;
    }

    String getManager() {
        return manager;
    }

    public Properties setManager(String manager) {
        if (manager != null && manager.length() > 65536) {
            throw new IllegalStateException("The length of manager must be less than or equal to 65536: " + manager.length());
        }
        this.manager = manager;
        return this;
    }

    String getCompany() {
        return company;
    }

    public Properties setCompany(String company) {
        if (company != null && company.length() > 65536) {
            throw new IllegalStateException("The length of company must be less than or equal to 65536: " + title.length());
        }
        this.company = company;
        return this;
    }

    String getHyperlinkBase() {
        return hyperlinkBase;
    }

    public Properties setHyperlinkBase(String hyperlinkBase) {
        if (hyperlinkBase != null && hyperlinkBase.length() > 65536) {
            throw new IllegalStateException("The length of hyperlinkBase must be less than or equal to 65536: " + title.length());
        }
        this.hyperlinkBase = hyperlinkBase;
        return this;
    }

    interface CustomProty<T> {
        void write(Writer w, int pid) throws IOException;
    }

    abstract class AbstractProperty<T> implements CustomProty<T> {
        protected String key;
        protected T value;

        public AbstractProperty(String key, T value) {
            this.key = key;
            this.value = value;
        }

        @Override
        public int hashCode() {
            return key.hashCode();
        }

        @Override
        public boolean equals(Object obj) {
            if (obj == null) {
                return false;
            }
            if (!(obj instanceof CustomProty)) {
                return false;
            }
            return this.key.equals(((AbstractProperty<?>) obj).key);
        }
    }

    class TextProperty extends AbstractProperty<String> {
        public TextProperty(String key, String value) {
            super(key, value);
        }

        @Override
        public void write(Writer w, int pid) throws IOException {
            w.append("<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"" + pid + "\" name=\"" + key + "\"><vt:lpwstr>" + value + "</vt:lpwstr></property>");
        }
    }

    class DateProperty extends AbstractProperty<Instant> {
        public DateProperty(String key, Instant value) {
            super(key, value);
        }

        @Override
        public void write(Writer w, int pid) throws IOException {
            w.append("<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"" + pid + "\" name=\"" + key + "\"><vt:filetime>" + DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSSX").withZone(ZoneId.of("UTC")).format(value) + "</vt:filetime></property>");
        }
    }

    class NumberProperty extends AbstractProperty<BigDecimal> {
        public NumberProperty(String key, BigDecimal value) {
            super(key, value);
        }

        @Override
        public void write(Writer w, int pid) throws IOException {
            w.append("<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"" + pid + "\" name=\"" + key + "\"><vt:r8>" + value.toPlainString() + "</vt:r8></property>");
        }
    }

    class BoolProperty extends AbstractProperty<Boolean> {
        public BoolProperty(String key, Boolean value) {
            super(key, value);
        }

        @Override
        public void write(Writer w, int pid) throws IOException {
            w.append("<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"" + pid + "\" name=\""+key+"\"><vt:bool>" + value.toString() + "</vt:bool></property>");
        }
    }

    public Properties setTextProperty(String key, String textValue) {
        customProperties.add(new TextProperty(key, textValue));
        return this;
    }

    public Properties setDateProperty(String key, Instant dateValue) {
        customProperties.add(new DateProperty(key, dateValue));
        return this;
    }

    public Properties setNumberProperty(String key, BigDecimal numberValue) {
        customProperties.add(new NumberProperty(key, numberValue));
        return this;
    }

    public Properties setBoolProperty(String key, Boolean boolValue) {
        customProperties.add(new BoolProperty(key, boolValue));
        return this;
    }

    public boolean hasCustomProperties() {
        return customProperties.size() > 0;
    }

    void writeCustomProperties(Writer w) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        w.append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
        Iterator<CustomProty> iterator = customProperties.iterator();
        for (int i = 0; iterator.hasNext(); i++) {
            int pid = i + 2;
            CustomProty next = iterator.next();
            next.write(w, pid);
        }
        w.append("</Properties>");
    }

}
