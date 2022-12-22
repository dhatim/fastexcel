package org.dhatim.fastexcel;

public class CustumPropertyValue {
    private String value;

    private Properties.CustomPropertyType type;

    private CustumPropertyValue(String value, Properties.CustomPropertyType type) {
        this.value = value;
        this.type = type;
    }

    public static CustumPropertyValue build( String value, Properties.CustomPropertyType type){
        return new CustumPropertyValue(value, type);
    }

    String getValue() {
        return value;
    }

    Properties.CustomPropertyType getType() {
        return type;
    }
}
