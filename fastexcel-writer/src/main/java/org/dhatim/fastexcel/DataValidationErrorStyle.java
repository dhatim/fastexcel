package org.dhatim.fastexcel;

/**
 * @author rjewing
 */
public enum DataValidationErrorStyle {
    STOP, WARNING, INFORMATION;

    @Override
    public String toString() {
        return this.name().toLowerCase();
    }
}
