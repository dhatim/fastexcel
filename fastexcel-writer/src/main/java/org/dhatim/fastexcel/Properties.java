package org.dhatim.fastexcel;

import java.util.*;

public class Properties {

    //****  core properties  ****
    private String title;
    private String subject;
    //alias:Tags
    private String keywords;
    //alias:Comments
    private String description;
    private String category;
    //****  app properties  ****
    private String manager;
    private String company;
    private String hyperlinkBase;

    private Map<String,CustumPropertyValue> customProperties = Collections.synchronizedMap(new HashMap<>());

    String getTitle() {
        return title;
    }

    public Properties setTitle(String title){
        if (title!=null&&title.length()>65536) {
            throw new IllegalStateException("The length of title must be less than or equal to 65536: "+title.length());
        }
        this.title = title;
        return this;
    }

    String getSubject() {
        return subject;
    }

    public Properties setSubject(String subject) {
        if (subject!=null&&subject.length()>65536) {
            throw new IllegalStateException("The length of subject must be less than or equal to 65536: "+subject.length());
        }
        this.subject = subject;
        return this;
    }

    String getKeywords() {
        return keywords;
    }

    public Properties setKeywords(String keywords) {
        if (keywords!=null&&keywords.length()>65536) {
            throw new IllegalStateException("The length of keywords must be less than or equal to 65536: "+keywords.length());
        }
        this.keywords = keywords;
        return this;
    }

    String getDescription() {
        return description;
    }

    public Properties setDescription(String description) {
        if (description!=null&&description.length()>65536) {
            throw new IllegalStateException("The length of description must be less than or equal to 65536: "+description.length());
        }
        this.description = description;
        return this;
    }

    String getCategory() {
        return category;
    }

    public Properties setCategory(String category) {
        if (category!=null&&category.length()>65536) {
            throw new IllegalStateException("The length of category must be less than or equal to 65536: "+category.length());
        }
        this.category = category;
        return this;
    }

    String getManager() {
        return manager;
    }

    public Properties setManager(String manager) {
        if (manager!=null&&manager.length()>65536) {
            throw new IllegalStateException("The length of manager must be less than or equal to 65536: "+manager.length());
        }
        this.manager = manager;
        return this;
    }

    String getCompany() {
        return company;
    }

    public Properties setCompany(String company) {
        if (company!=null&&company.length()>65536) {
            throw new IllegalStateException("The length of company must be less than or equal to 65536: "+title.length());
        }
        this.company = company;
        return this;
    }

    String getHyperlinkBase() {
        return hyperlinkBase;
    }

    public Properties setHyperlinkBase(String hyperlinkBase) {
        if (hyperlinkBase!=null&&hyperlinkBase.length()>65536) {
            throw new IllegalStateException("The length of hyperlinkBase must be less than or equal to 65536: "+title.length());
        }
        this.hyperlinkBase = hyperlinkBase;
        return this;
    }

    public Properties setCustomProperties(String key,String value, CustomPropertyType type){
        customProperties.put(key,CustumPropertyValue.build(value,type));
        return this;
    }

    public Map<String, CustumPropertyValue> getCustomProperties() {
        return customProperties;
    }

    public enum CustomPropertyType{
        TEXT,DATE,NUMBER,YES_OR_NO
    }
}
