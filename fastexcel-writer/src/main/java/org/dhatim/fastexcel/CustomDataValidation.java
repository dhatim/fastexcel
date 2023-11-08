package org.dhatim.fastexcel;

import java.io.IOException;

/**
 * A CustomDataValidation defines a DataValidation for a worksheet of type = "custom"
 */
public class CustomDataValidation implements DataValidation {
    private final static String TYPE = "custom";
    private final Range range;
    private final Formula formula;

    private boolean allowBlank = true;
    private boolean showDropdown = true;
    private DataValidationErrorStyle errorStyle = DataValidationErrorStyle.INFORMATION;
    private boolean showErrorMessage = false;
    private String errorTitle;
    private String error;

    /**
     * Constructor
     *
     * @param range     The Range this validation is applied to
     * @param formula   The Formula of the custom validation
     */
    CustomDataValidation(Range range, Formula formula) {
        this.range = range;
        this.formula = formula;
    }

    /**
     * whether blank cells should pass the validation
     *
     * @param allowBlank whether or not to allow blank values
     * @return this CustomDataValidation
     */
    public CustomDataValidation allowBlank(boolean allowBlank) {
        this.allowBlank = allowBlank;
        return this;
    }

    /**
     * Whether Excel will show an in-cell dropdown list
     * containing the validation list
     *
     * @param showDropdown whether or not to show the dropdown
     * @return this CustomDataValidation
     */
    public CustomDataValidation showDropdown(boolean showDropdown) {
        this.showDropdown = showDropdown;
        return this;
    }

    /**
     * The style of error alert used for this data validation.
     *
     * @param errorStyle The DataValidationErrorStyle for this DataValidation
     * @return this CustomDataValidation
     */
    public CustomDataValidation errorStyle(DataValidationErrorStyle errorStyle) {
        this.errorStyle = errorStyle;
        return this;
    }

    /**
     * Whether to display the error alert message when an invalid value has been entered.
     *
     * @param showErrorMessage whether to display the error message
     * @return this CustomDataValidation
     */
    public CustomDataValidation showErrorMessage(boolean showErrorMessage) {
        this.showErrorMessage = showErrorMessage;
        return this;
    }

    /**
     * Title bar text of error alert.
     *
     * @param errorTitle The error title
     * @return this CustomDataValidation
     */
    public CustomDataValidation errorTitle(String errorTitle) {
        this.errorTitle = errorTitle;
        return this;
    }

    /**
     * Message text of error alert.
     *
     * @param error The error message
     * @return this CustomDataValidation
     */
    public CustomDataValidation error(String error) {
        this.error = error;
        return this;
    }

    /**
     * Write this dataValidation as an XML element.
     *
     * @param w Output writer.
     * @throws IOException If an I/O error occurs.
     */
    @Override
    public void write(Writer w) throws IOException {
        w
                .append("<dataValidation sqref=\"")
                .append(range.toString())
                .append("\" type=\"")
                .append(TYPE)
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
                .append("\"><formula1>")
                .append(formula.getExpression())
                .append("</formula1></dataValidation>");
    }
}
