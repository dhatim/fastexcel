package org.dhatim.fastexcel;

/**
 * This simple type defines the possible states for sheet visibility.
 *
 * This simple type's contents are a restriction of the XML Schema string datatype.
 */
public enum VisibilityState {

    /**
     * Indicates the book window is hidden, but can be shown by the user via the user interface.
     */
    HIDDEN("hidden"),

    /**
     * Indicates the sheet is hidden and cannot be shown in the user interface (UI). This state is only available
     * programmatically.
     */
    VERY_HIDDEN("veryHidden"),

    /**
     * Indicates the sheet is visible.
     */
    VISIBLE("visible");

    private final String name;

    /**
     * Constructor that sets the name.
     *
     * @param name The name of the visibility state.
     */
    VisibilityState(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }
}
