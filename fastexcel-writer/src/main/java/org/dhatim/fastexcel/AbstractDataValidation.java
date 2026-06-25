package org.dhatim.fastexcel;

import java.io.IOException;

abstract class AbstractDataValidation<T extends AbstractDataValidation<T>> implements DataValidation {

    private final Range range;
    private final String type;

    private boolean allowBlank = true;
    private boolean showDropdown = true;
    private DataValidationErrorStyle errorStyle = DataValidationErrorStyle.INFORMATION;
    private boolean showErrorMessage = false;
    private String errorTitle;
    private String error;

    AbstractDataValidation(Range range, String type) {
        this.range = range;
        this.type = type;
    }

    public T allowBlank(boolean allowBlank) {
        this.allowBlank = allowBlank;
        return self();
    }

    public T showDropdown(boolean showDropdown) {
        this.showDropdown = showDropdown;
        return self();
    }

    public T errorStyle(DataValidationErrorStyle errorStyle) {
        this.errorStyle = errorStyle;
        return self();
    }

    public T showErrorMessage(boolean showErrorMessage) {
        this.showErrorMessage = showErrorMessage;
        return self();
    }

    public T errorTitle(String errorTitle) {
        this.errorTitle = errorTitle;
        return self();
    }

    public T error(String error) {
        this.error = error;
        return self();
    }

    @Override
    public final void write(Writer w) throws IOException {
        w.append("<dataValidation sqref=\"")
                .append(range.toString())
                .append("\" type=\"")
                .append(type)
                .append("\" allowBlank=\"")
                .append(String.valueOf(allowBlank))
                .append("\" showDropDown=\"")
                .append(String.valueOf(!showDropdown)) // for some reason, this is the inverse of what you'd expect
                .append("\" errorStyle=\"")
                .append(errorStyle.toString())
                .append("\" showErrorMessage=\"")
                .append(String.valueOf(showErrorMessage))
                .append("\" errorTitle=\"")
                .append(errorTitle)
                .append("\" error=\"")
                .append(error)
                .append("\"><formula1>");
        writeFormula(w);
        w.append("</formula1></dataValidation>");
    }

    protected abstract void writeFormula(Writer w) throws IOException;

    @SuppressWarnings("unchecked")
    private T self() {
        return (T) this;
    }
}
