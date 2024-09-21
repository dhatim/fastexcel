/*
 * Copyright 2016 Dhatim.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
package org.dhatim.fastexcel;

import java.io.Closeable;
import java.io.IOException;
import java.util.*;

public abstract class AbstractWorksheet implements Closeable {

    /**
     * Maximum number of rows in Excel.
     */
    public static final int MAX_ROWS = 1_048_576;

    /**
     * Maximum number of columns in Excel.
     */
    public static final int MAX_COLS = 16_384;

    /**
     * Maximum column width in Excel.
     */
    public static final int MAX_COL_WIDTH = 255;

    /**
     * Default column width in Excel.
     */
    public static final double DEFAULT_COL_WIDTH = 8.88671875;

    /**
     * Maximum row height in Excel.
     */
    public static final double MAX_ROW_HEIGHT = 409.5;

    /**
     * Parent workbook.
     */
    protected final Workbook workbook;

    /**
     * Worksheet name.
     */
    protected final String name;

    /**
     * The Writer instance.
     */
    protected Writer writer;

    /**
     * The visibility state of this sheet.
     */
    private VisibilityState visibilityState;

    /**
     * Is this worksheet construction completed?
     */
    protected boolean finished;

    /**
     * The set of protection options that are applied on the sheet.
     */
    protected Set<SheetProtectionOption> sheetProtectionOptions;

    /**
     * The hashed password that protects this sheet.
     */
    protected String passwordHash;

    /**
     * The tab RGB color.
     */
    protected String tabColor;

    /**
     * Whether grid lines are displayed
     */
    private boolean showGridLines = true;

    /**
     * Display the worksheet from right to left
     */
    private boolean rightToLeft = false;

    /**
     * Flag indicating whether summary rows appear below detail in an outline, when applying an outline.
     */
    private boolean rowSumsBelow = true;

    /**
     * Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline.
     */
    private boolean rowSumsRight = true;

    /**
     * Sheet view zoom percentage
     */
    private int zoomScale = 100;

    /**
     * Number of top rows that will remain frozen while scrolling.
     */
    private int freezeTopRows = 0;

    /**
     * Number of columns from the left that remain frozen while scrolling.
     */
    private int freezeLeftColumns = 0;

    /**
     * Auto page breaks.
     */
    protected Boolean autoPageBreaks = false;

    /**
     * Fit to page (true for fit to width/height).
     */
    protected Boolean fitToPage = false;

    /**
     * Fit to width in the print setup.
     */
    protected int fitToWidth = 1;

    /**
     * Fit to height in the print setup.
     */
    protected int fitToHeight = 1;

    /**
     * List of conditional formattings for this worksheet
     */
    protected final List<ConditionalFormatting> conditionalFormattings = new ArrayList<>();

    /**
     * Header map for left, central and right field text.
     */
    private final Map<Position, MarginalInformation> header = new LinkedHashMap<>();

    /**
     * Footer map for left, central and right field text.
     */
    private final Map<Position, MarginalInformation> footer = new LinkedHashMap<>();

    /**
     * Page orientation [landscape / portrait] for the print preview setup.
     */
    private String pageOrientation = "portrait";

    /**
     * Paper size for the print preview setup.
     */
    private PaperSize paperSize = PaperSize.LETTER_PAPER;

    /**
     * Scaling factor for the print setup.
     */
    private int pageScale = 100;

    /**
     * First page number in the print setup.
     */
    private int firstPageNumber = 0;

    /**
     * Whether to use the firstPageNumber in the print setup.
     */
    private Boolean useFirstPageNumber = false;

    /**
     * Black and white mode in the print setup.
     */
    private Boolean blackAndWhite = false;

    /**
     * Header margin value in inches.
     */
    private float headerMargin = 0.3f;

    /**
     * Footer margin value in inches.
     */
    private float footerMargin = 0.3f;

    /**
     * Top margin value in inches.
     */
    private float topMargin = 0.75f;

    /**
     * Bottom margin value in inches.
     */
    private float bottomMargin = 0.75f;

    /**
     * Left margin value in inches.
     */
    private float leftMargin = 0.7f;

    /**
     * Right margin value in inches.
     */
    protected float rightMargin = 0.7f;

    protected final Comments comments = new Comments();

    protected final Map<String,Table> tables = new LinkedHashMap<>();

    private final DynamicBitMatrix tablesMatrix = new DynamicBitMatrix(MAX_COLS, MAX_ROWS);

    private Relationships relationships;

    /**
     * List of DataValidations for this worksheet
     */
    private final List<DataValidation> dataValidations = new ArrayList<>();

    private Map<HyperLink, Ref> hyperlinkRanges = new LinkedHashMap<>();

    /**
     * Constructor.
     *
     * @param workbook Parent workbook.
     * @param name     Worksheet name.
     */
    AbstractWorksheet(Workbook workbook, String name) {
        this.workbook = Objects.requireNonNull(workbook);
        this.name = Objects.requireNonNull(name);
        this.relationships = new Relationships(this);
    }

    @Override
    public void close() throws IOException {
        finish();
    }

    public abstract void initDocumentAndFlush() throws IOException;

    protected Writer initWriter() throws IOException {
        int index = workbook.getIndex(this);
        writer = workbook.beginFile("xl/worksheets/sheet" + index + ".xml");
        return writer;
    }

    /**
     * Finish the construction of this worksheet. This creates the worksheet
     * file on the workbook's output stream. Rows and cells in this worksheet
     * are then destroyed.
     *
     * @throws IOException If an I/O error occurs.
     */
    public abstract void finish() throws IOException;

    /**
     * Get worksheet name.
     *
     * @return Worksheet name.
     */
    public String getName() {
        return name;
    }

    /**
     * Get parent workbook.
     *
     * @return Parent workbook.
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Sets the visibility state of the sheet
     * <p>
     * This is done by setting the {@code state} attribute in the workbook.xml.
     *
     * @param visibilityState New visibility state for this sheet.
     */
    public void setVisibilityState(VisibilityState visibilityState) {
        this.visibilityState = visibilityState;
    }

    public VisibilityState getVisibilityState() {
        return visibilityState;
    }

    /**
     * @return True if this worksheet construction is completed.
     */
    public boolean isFinished() {
        return finished;
    }

    /**
     * Protects the sheet with a password. This method protects all the default {@link SheetProtectionOption}s and
     * 'sheet'. (Note that this is not very secure and only meant for discouraging changes.)
     *
     * @param password The password to use.
     */
    public void protect(String password) {
        protect(password, SheetProtectionOption.DEFAULT_OPTIONS);
    }

    /**
     * Protects the sheet with a password. (Note that this is not very secure and only meant for discouraging changes.)
     *
     * @param password The password to use.
     * @param options  An array of all the {@link SheetProtectionOption}s to protect.
     */
    public void protect(String password, SheetProtectionOption... options) {
        final EnumSet<SheetProtectionOption> optionSet = EnumSet.noneOf(SheetProtectionOption.class);
        Collections.addAll(optionSet, options);
        protect(password, optionSet);
    }

    /**
     * Protects the sheet with a password. (Note that this is not very secure and only meant for discouraging changes.)
     *
     * @param password The password to use.
     * @param options  A {@link Set} of all the {@link SheetProtectionOption}s to protect.
     */
    public void protect(String password, Set<SheetProtectionOption> options) {
        if (password == null) {
            this.sheetProtectionOptions = null;
            this.passwordHash = null;
            return;
        }
        this.sheetProtectionOptions = options;
        this.passwordHash = hashPassword(password);
    }

    public void initDocumentTags() throws IOException {
        writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        writer.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        writer.append("<sheetPr filterMode=\"" + "false" + "\">");
        if (tabColor != null) {
            writer.append("<tabColor rgb=\"" + tabColor + "\"/>");
        }
        if (!rowSumsBelow || !rowSumsRight) {
            writer.append("<outlinePr ");
            if (!rowSumsBelow) {
                writer.append("summaryBelow=\"0\" ");
            }
            if (!rowSumsRight) {
                writer.append("summaryRight=\"0\" ");
            }
            writer.append("/>");
        }
        writer.append("<pageSetUpPr fitToPage=\"" + fitToPage + "\" " + "autoPageBreaks=\"" + autoPageBreaks + "\"/></sheetPr>");
        writer.append("<dimension ref=\"A1\"/>");
        writer.append("<sheetViews><sheetView workbookViewId=\"0\"");
        if (!showGridLines) {
            writer.append(" showGridLines=\"false\"");
        }
        if (rightToLeft) {
            writer.append(" rightToLeft=\"true\"");
        }
        if (zoomScale != 100) {
            writer.append(" zoomScale=\"").append(zoomScale).append("\"");
        }
        writer.append(">");
        if (freezeLeftColumns > 0 || freezeTopRows > 0) {
            writeFreezePane(writer);
        }
        writer.append("</sheetView>");
        writer.append("</sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/>");
    }

    /**
     * Writes corresponding pane definitions into XML and freezes pane.
     */
    protected void writeFreezePane(Writer w) throws IOException {
        String activePane = freezeLeftColumns == 0 ? "bottomLeft" : freezeTopRows == 0 ? "topRight" : "bottomRight";
        String freezePane = "<pane xSplit=\"" + freezeLeftColumns +
                "\" ySplit=\"" + freezeTopRows + "\" topLeftCell=\"" +
                getCellMark(freezeTopRows, freezeLeftColumns) +
                "\" activePane=\"" + activePane + "\" state=\"frozen\"/>";
        w.append(freezePane);
        String topLeftPane = "<selection pane=\"topLeft\" activeCell=\"" +
                getCellMark(0, 0) +
                "\" activeCellId=\"0\" sqref=\"" +
                getCellMark(0, 0) + "\"/>";
        w.append(topLeftPane);
        if (freezeLeftColumns != 0) {
            String topRightPane = "<selection pane=\"topRight\" activeCell=\"" +
                    getCellMark(0, freezeLeftColumns) +
                    "\" activeCellId=\"0\" sqref=\"" +
                    getCellMark(0, freezeLeftColumns) + "\"/>";
            w.append(topRightPane);
        }
        if (freezeTopRows != 0) {
            String bottomLeftPane = "<selection pane=\"bottomLeft\" activeCell=\"" +
                    getCellMark(freezeTopRows, 0) +
                    "\" activeCellId=\"0\" sqref=\"" +
                    getCellMark(freezeTopRows, 0) + "\"/>";
            w.append(bottomLeftPane);
        }
        if (freezeLeftColumns != 0 && freezeTopRows != 0) {
            String bottomRightPane = "<selection pane=\"bottomRight\" activeCell=\"" +
                    getCellMark(freezeTopRows, freezeLeftColumns) +
                    "\" activeCellId=\"0\" sqref=\"" +
                    getCellMark(freezeTopRows, freezeLeftColumns) + "\"/>";
            w.append(bottomRightPane);
        }
    }

    protected void setupSheetPassword(Writer writer) throws IOException {
        if (passwordHash != null) {
            writer.append("<sheetProtection password=\"").append(passwordHash).append("\" ");
            for (SheetProtectionOption option : SheetProtectionOption.values()) {
                if (option.getDefaultValue() != sheetProtectionOptions.contains(option)) {
                    writer.append(option.getName()).append("=\"").append(Boolean.toString(!option.getDefaultValue())).append("\" ");
                }
            }
            writer.append("/>");
        }
    }

    /**
     * @param rgbColor FFF381E0
     */
    public void setTabColor(String rgbColor) {
        this.tabColor = rgbColor;
    }

    public void setAutoPageBreaks(Boolean autoPageBreaks) {
        this.autoPageBreaks = autoPageBreaks;
    }

    public void setFitToPage(Boolean fitToPage) {
        this.fitToPage = fitToPage;
    }

    public void fitToHeight(Short fitToHeight) {
        this.fitToPage = true;
        this.fitToHeight = fitToHeight;
    }

    public void fitToWidth(Short fitToWidth) {
        this.fitToPage = true;
        this.fitToWidth = fitToWidth;
    }

    /**
     * Hide grid lines.
     */
    public void hideGridLines() {
        this.showGridLines = false;
    }

    /**
     * Display the worksheet from right to left
     */
    public void rightToLeft() {
        this.rightToLeft = true;
    }


    /**
     * Set sheet view zoom level in percent. Default is 100 (100%).
     *
     * @param zoomPercent - zoom level from 10 to 400
     */
    public void setZoom(int zoomPercent) {
        if (10 <= zoomPercent && zoomPercent <= 400) {
            this.zoomScale = zoomPercent;
        } else {
            throw new IllegalArgumentException("zoom must be within 10 and 400 inclusive");
        }
    }

    /**
     * Set freeze pane (rows and columns that remain when scrolling).
     *
     * @param nLeftColumns - number of columns from the left that will remain frozen
     * @param nTopRows     - number of rows from the top that will remain frozen
     */
    public void freezePane(int nLeftColumns, int nTopRows) {
        this.freezeLeftColumns = nLeftColumns;
        this.freezeTopRows = nTopRows;
    }

    /**
     * Unfreeze any frozen rows, or columns.
     */
    public void unfreeze() {
        this.freezeLeftColumns = 0;
        this.freezeTopRows = 0;
    }

    public void rowSumsBelow(boolean rowSumsBelow) {
        this.rowSumsBelow = rowSumsBelow;
    }

    public void rowSumsRight(boolean rowSumsRight) {
        this.rowSumsRight = rowSumsRight;
    }

    void addConditionalFormatting(ConditionalFormatting conditionalFormatting) {
        conditionalFormattings.add(conditionalFormatting);
    }

    /**
     * Set header margin.
     *
     * @param margin - header margin in inches
     */
    public void headerMargin(float margin) {
        this.headerMargin = margin;
    }

    /**
     * Set footer margin.
     *
     * @param margin - footer page margin in inches
     */
    public void footerMargin(float margin) {
        this.footerMargin = margin;
    }

    /**
     * Set top margin.
     *
     * @param margin - top page margin in inches
     */
    public void topMargin(float margin) {
        this.topMargin = margin;
    }

    /**
     * Set bottom margin.
     *
     * @param margin - bottom page margin in inches
     */
    public void bottomMargin(float margin) {
        this.bottomMargin = margin;
    }

    /**
     * Set left margin.
     *
     * @param margin - left page margin in inches
     */
    public void leftMargin(float margin) {
        this.leftMargin = margin;
    }

    /**
     * Set right margin.
     *
     * @param margin - right page margin in inches
     */
    public void rightMargin(float margin) {
        this.rightMargin = margin;
    }

    /**
     * Set the page orientation.
     *
     * @param orientation New page orientation for this worksheet
     */
    public void pageOrientation(String orientation) {
        this.pageOrientation = orientation;
    }

    /**
     * Set the paper size.
     *
     * @param size New paper size for this worksheet
     */
    public void paperSize(PaperSize size) {
        this.paperSize = size;
    }

    /**
     * @param scale = scaling factor for the print setup (between 1 and 100)
     */
    public void pageScale(int scale) {
        this.pageScale = scale;
    }

    /**
     * @param pageNumber - first page number (default: 0)
     */
    public void firstPageNumber(int pageNumber) {
        this.firstPageNumber = pageNumber;
        this.useFirstPageNumber = true;
    }

    public void printInBlackAndWhite() {
        this.blackAndWhite = true;
    }

    public void printInColor() {
        this.blackAndWhite = false;
    }

    /**
     * Set footer text.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     */
    public void footer(String text, Position position) {
        this.footer.put(position, new MarginalInformation(text, position));
    }

    /**
     * Set footer text with specified font size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontSize - integer describing font size
     */
    public void footer(String text, Position position, int fontSize) {
        this.footer.put(position, new MarginalInformation(text, position)
                .withFontSize(fontSize));
    }

    /**
     * Set footer text with specified font and size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontName - font name (e.g., "Arial")
     * @param fontSize - integer describing font size
     */
    public void footer(String text, Position position, String fontName, int fontSize) {
        this.footer.put(position, new MarginalInformation(text, position)
                .withFont(fontName)
                .withFontSize(fontSize));
    }

    /**
     * Set header text.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontName - font name (e.g., "Arial")
     * @param fontSize - integer describing font size
     */
    public void header(String text, Position position, String fontName, int fontSize) {
        this.header.put(position, new MarginalInformation(text, position)
                .withFont(fontName)
                .withFontSize(fontSize));
    }

    /**
     * Set header text with specified font size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     * @param fontSize - integer describing font size
     */
    public void header(String text, Position position, int fontSize) {
        this.header.put(position, new MarginalInformation(text, position)
                .withFontSize(fontSize));
    }

    /**
     * Set header text with specified font and size.
     * @param text - text input form or custom text
     * @param position - Position.LEFT/RIGHT/CENTER enum
     */
    public void header(String text, Position position) {
        this.header.put(position, new MarginalInformation(text, position));
    }

    protected void setupDataValidations() throws IOException {
        if (!dataValidations.isEmpty()) {
            writer.append("<dataValidations count=\"").append(dataValidations.size()).append("\">");
            for (DataValidation v: dataValidations) {
                v.write(writer);
            }
            writer.append("</dataValidations>");
        }
    }

    protected void setupHyperlinkRanges() throws IOException {
        if (!hyperlinkRanges.isEmpty()) {
            writer.append("<hyperlinks>");
            for (Map.Entry<HyperLink, Ref> hr : hyperlinkRanges.entrySet()) {
                HyperLink hyperLink = hr.getKey();
                writer.append("<hyperlink ");
                Ref ref = hr.getValue();
                writer.append("ref=\"" + ref.toString()+"\" ");
                if (hyperLink.getHyperLinkType().equals(HyperLinkType.EXTERNAL)) {
                    String rId = relationships.setHyperLinkRels(hyperLink.getLinkStr(), "External");
                    writer.append("r:id=\"" + rId +"\" ");
                }else{
                    writer.append("location=\"").append(hyperLink.getLinkStr()).append("\"");
                }
                writer.append("/>");
            }
            writer.append("</hyperlinks>");
        }
    }

    protected void setupPageMargins() throws IOException {
        /* set page margins for the print setup (see in print preview) */
        String margins = "<pageMargins bottom=\"" + bottomMargin +
                "\" footer=\"" + footerMargin +
                "\" header=\"" + headerMargin +
                "\" left=\"" + leftMargin +
                "\" right=\"" + rightMargin +
                "\" top=\"" + topMargin + "\"/>";
        writer.append(margins);
    }

    protected void setupPage() throws IOException {
        /* set page orientation for the print setup */
        writer.append("<pageSetup")
                .append(" paperSize=\"" + paperSize.xmlValue + "\"")
                .append(" scale=\"" + pageScale + "\"")
                .append(" fitToWidth=\"" + fitToWidth + "\"")
                .append(" fitToHeight=\"" + fitToHeight + "\"")
                .append(" firstPageNumber=\"" + firstPageNumber + "\"")
                .append(" useFirstPageNumber=\"" + useFirstPageNumber.toString() + "\"")
                .append(" blackAndWhite=\"" + blackAndWhite.toString() + "\"")
                .append(" orientation=\"" + pageOrientation + "\"")
                .append("/>");
    }

    protected void setupFooter() throws IOException {
        /* write to header and footer */
        writer.append("<headerFooter differentFirst=\"false\" differentOddEven=\"false\">");
        writer.append("<oddHeader>");
        for (MarginalInformation headerEntry : header.values()) {
            headerEntry.write(writer);
        }
        writer.append("</oddHeader>");
        writer.append("<oddFooter>");
        for (MarginalInformation footerEntry : footer.values()) {
            footerEntry.write(writer);
        }
        writer.append("</oddFooter></headerFooter>");
    }

    protected void setupComments() throws IOException {
        if (!comments.isEmpty()) {
            writer.append("<drawing r:id=\"d\"/>");
            writer.append("<legacyDrawing r:id=\"v\"/>");
        }
    }

    protected void writeComments(int index) throws IOException {
        /* write comment files */
        if (!comments.isEmpty()) {
            workbook.writeFile("xl/comments" + index + ".xml", comments::writeComments);
            workbook.writeFile("xl/drawings/vmlDrawing" + index + ".vml", comments::writeVmlDrawing);
            workbook.writeFile("xl/drawings/drawing" + index + ".xml", comments::writeDrawing);
            relationships.setCommentsRels(index);
        }
    }

    /**
     * Assign a note/comment to a cell.
     * The comment popup will be twice the size of the cell and will be initially hidden.
     * <p>
     * Comments are stored in memory till call to {@link #close()} (or  the old fashion way {@link #finish()}) - calling {@link #initDocumentAndFlush()} does not write them to output stream.
     * @param r Zero-based row number.
     * @param c Zero-based column number.
     * @param comment Note text
     */
    public void comment(int r, int c, String comment) {
        comments.set(r, c, comment);
    }

    protected void setupTables() throws IOException {
        if (!tables.isEmpty()) {
            writer.append("<tableParts count=\"" + tables.size() + "\">");
            for (Map.Entry<String, Table> entry : tables.entrySet()) {
                writer.append("<tablePart r:id=\"" + entry.getKey() + "\"/>");
            }
            writer.append("</tableParts>");
        }
    }

    protected void writeTables() throws IOException {
        // write table files
        for (Map.Entry<String, Table> entry : tables.entrySet()) {
            Table table = entry.getValue();
            workbook.writeFile("xl/tables/table" + table.index + ".xml",table::write);
        }
    }

    Table addTable(Range range, String... headers) {
        if (!tablesMatrix.isConflict(range.getTop(), range.getLeft(), range.getBottom(), range.getRight())) {
            int tableIndex = getWorkbook().nextTableIndex();
            String rId = relationships.setTableRels(tableIndex);
            Table table = new Table(tableIndex, range, headers);
            tables.put(rId, table);
            tablesMatrix.setRegion(range.getTop(), range.getLeft(), range.getBottom(), range.getRight());
            return table;
        } else {
            throw new IllegalArgumentException("Table conflicted:" + range);
        }
    }

    protected void writeRelationships(int index) throws IOException {
        // write relationship files
        if (!relationships.isEmpty()) {
            workbook.writeFile("xl/worksheets/_rels/sheet"+ index +".xml.rels",relationships::write);
        }
    }

    void addValidation(DataValidation validation) {
        dataValidations.add(validation);
    }

    void addHyperlink(Ref ref, HyperLink hyperLink) {
        this.hyperlinkRanges.put(hyperLink, ref);
    }

    /**
     * Hash the password.
     *
     * @param password The password to hash.
     * @return The password hash as a hex string (2 bytes)
     */
    static String hashPassword(String password) {
        byte[] passwordCharacters = password.getBytes();
        int hash = 0;
        if (passwordCharacters.length > 0) {
            int charIndex = passwordCharacters.length;
            while (charIndex-- > 0) {
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters[charIndex];
            }
            // also hash with char-count
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
            hash ^= passwordCharacters.length;
            hash ^= (0x8000 | ('N' << 8) | 'K');
        }

        return Integer.toHexString(hash & 0xffff);
    }

    /**
     * Helper method to get a cell name from (x, y) cell position.
     * e.g. "B3" from cell position (2, 1)
     */
    static String getCellMark(int row, int coll) {
        return ((char) ('A' + coll)) + String.valueOf(row + 1);
    }

    /**
     * Write a column as an XML element.
     *
     * @param w           Output writer.
     * @param columnIndex Zero-based column number.
     * @param maxWidth    The maximum width
     * @param bestFit     Whether or not this column should be optimized for fit
     * @param isHidden    Whether or not this column is hidden
     * @param style       Cached style index of the column
     * @throws IOException If an I/O error occurs.
     */
    static void writeCol(Writer w, int columnIndex, double maxWidth, boolean bestFit, boolean isHidden, int groupLevel,
                         int style) throws IOException {
        final int col = columnIndex + 1;
        w.append("<col min=\"").append(col).append("\" max=\"").append(col).append("\" width=\"")
                .append(Math.min(MAX_COL_WIDTH, maxWidth));
        w.append("\" outlineLevel=\"").append(groupLevel);
        w.append("\" customWidth=\"true\" bestFit=\"")
                .append(String.valueOf(bestFit));

        if (isHidden) {
            w.append("\" hidden=\"true");
        }
        w.append("\"");
        if (style > 0) {
            w.append(" style=\"").append(style).append("\"");
        }

        w.append("/>");
    }

    /**
     * Write a row as an XML element.
     *
     * @param w          Output writer.
     * @param r          Zero-based row number.
     * @param isHidden   Whether or not this row is hidden
     * @param groupLevel Group level of row
     * @param rowHeight  Row height value in points to be set if customHeight is true
     * @param row        Cells in the row.
     * @throws IOException If an I/O error occurs.
     */
    static void writeRow(Writer w, int r, boolean isHidden, byte groupLevel,
                         Double rowHeight, Cell... row) throws IOException {
        w.append("<row r=\"").append(r + 1).append("\"");
        if (isHidden) {
            w.append(" hidden=\"true\"");
        }
        if (rowHeight != null) {
            w.append(" ht=\"")
                    .append(rowHeight)
                    .append("\"")
                    .append(" customHeight=\"1\"");
        }
        if (groupLevel != 0) {
            w.append(" outlineLevel=\"")
                    .append(groupLevel)
                    .append("\"");
        }
        w.append(">");
        if (null != row) {
            for (int c = 0; c < row.length; ++c) {
                if (row[c] != null) {
                    row[c].write(w, r, c);
                }
            }
        }
        w.append("</row>");
    }
}
