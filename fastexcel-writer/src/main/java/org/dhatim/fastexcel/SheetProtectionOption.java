package org.dhatim.fastexcel;

import java.util.EnumSet;
import java.util.Set;

/**
 * Represents an attribute on the &lt;sheetProtection&gt; xml-tag.
 */
public enum SheetProtectionOption {

    /**
     * Autofilters are locked when the sheet is protected.
     */
    AUTO_FILTER("autoFilter", true),

    /**
     * Deleting columns is locked when the sheet is protected.
     */
    DELETE_COLUMNS("deleteColumns", true),

    /**
     * Deleting rows is locked when the sheet is protected.
     */
    DELETE_ROWS("deleteRows", true),

    /**
     * Formatting cells is locked when the sheet is protected.
     */
    FORMAT_CELLS("formatCells", true),

    /**
     * Formatting columns is locked when the sheet is protected.
     */
    FORMAT_COLUMNS("formatColumns", true),

    /**
     * Formatting rows is locked when the sheet is protected.
     */
    FORMAT_ROWS("formatRows", true),

    /**
     * Inserting columns is locked when the sheet is protected.
     */
    INSERT_COLUMNS("insertColumns", true),

    /**
     * Inserting hyperlinks is locked when the sheet is protected.
     */
    INSERT_HYPERLINKS("insertHyperlinks", true),

    /**
     * Inserting rows is locked when the sheet is protected.
     */
    INSERT_ROWS("insertRows", true),

    /**
     * Pivot tables are locked when the sheet is protected.
     */
    PIVOT_TABLES("pivotTables", true),

    /**
     * Sorting is locked when the sheet is protected.
     */
    SORT("sort", true),

    /**
     * Sheet is locked when the sheet is protected.
     */
    SHEET("sheet", false),

    /**
     * Objects are locked when the sheet is protected.
     */
    OBJECTS("objects", false),

    /**
     * Scenarios are locked when the sheet is protected.
     */
    SCENARIOS("scenarios", false),

    /**
     * Selection of locked cells is locked when the sheet is protected.
     */
    SELECT_LOCKED_CELLS("selectLockedCells", false),

    /**
     * Selection of unlocked cells is locked when the sheet is protected.
     */
    SELECT_UNLOCKED_CELLS("selectUnlockedCells", false);

    /**
     * The options that are {@code true} by default AND 'sheet'.
     */
    public static final Set<SheetProtectionOption> DEFAULT_OPTIONS = EnumSet.range(AUTO_FILTER, SHEET);

    private final String name;

    private final boolean defaultValue;

    /**
     * Constructor that sets the name.
     *
     * @param name The name of the sheet protection option.
     */
    SheetProtectionOption(String name, boolean defaultValue) {
        this.name = name;
        this.defaultValue = defaultValue;
    }

    public String getName() {
        return name;
    }

    public boolean getDefaultValue() {
        return defaultValue;
    }}
