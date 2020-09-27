package com.github.javaxcel.styler;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;

import java.util.stream.IntStream;

public final class ExcelStyler {

    public static final int HSSF_MAX_ROWS = 65_536; // 2^16
    public static final int HSSF_MAX_COLUMNS = 256; // 2^8
    public static final int XSSF_MAX_ROWS = 1_048_576; // 2^20
    public static final int XSSF_MAX_COLUMNS = 16_384; // 2^14

    private ExcelStyler() {
    }

    /**
     * Adjusts width of columns to fit the contents.
     *
     * <p> This can be affected by font size and font family.
     * If you want this process well, set up the same font family into all cells.
     * This process will be perform in parallel.
     *
     * @param sheet        excel sheet
     * @param numOfColumns number of the columns that wanted to make fit contents.
     * @see Sheet#autoSizeColumn(int)
     */
    public static void autoResizeColumns(Sheet sheet, int numOfColumns) {
        IntStream.range(0, numOfColumns).parallel().forEach(sheet::autoSizeColumn);
    }

    /**
     * Hides extraneous rows.
     *
     * @param sheet     excel sheet
     * @param numOfRows number of the rows that have contents.
     * @see Row#setZeroHeight(boolean)
     */
    public static void hideExtraRows(Sheet sheet, int numOfRows) {
        int maxRows = sheet instanceof HSSFSheet
                ? HSSF_MAX_ROWS
                : XSSF_MAX_ROWS;

        for (int i = numOfRows; i < maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null) row = sheet.createRow(i);

            row.setZeroHeight(true);
        }
    }

    /**
     * Hides extraneous columns.
     *
     * @param sheet        excel sheet
     * @param numOfColumns number of the columns that have contents.
     * @see Sheet#setColumnHidden(int, boolean)
     */
    public static void hideExtraColumns(Sheet sheet, int numOfColumns) {
        int maxColumns = sheet instanceof HSSFSheet
                ? HSSF_MAX_COLUMNS
                : XSSF_MAX_COLUMNS;

        for (int i = numOfColumns; i < maxColumns; i++) {
            sheet.setColumnHidden(i, true);
        }
    }

    public static CellStyle drawBorder(CellStyle style) {
        BorderStyle border = BorderStyle.THIN;
        IndexedColors color = IndexedColors.BLACK;
        drawBorder(style, border, color);

        return style;
    }

    public static CellStyle drawBorder(CellStyle style, BorderStyle border, IndexedColors color) {
        style.setBorderTop(border);
        style.setBorderRight(border);
        style.setBorderBottom(border);
        style.setBorderLeft(border);

        style.setTopBorderColor(color.getIndex());
        style.setRightBorderColor(color.getIndex());
        style.setBottomBorderColor(color.getIndex());
        style.setLeftBorderColor(color.getIndex());

        return style;
    }

    public static CellStyle dyeBackground(CellStyle style) {
        dyeBackground(style, FillPatternType.SOLID_FOREGROUND, IndexedColors.GREY_25_PERCENT);
        return style;
    }

    public static CellStyle dyeBackground(CellStyle style, FillPatternType pattern, IndexedColors color) {
        style.setFillPattern(pattern);
        style.setFillForegroundColor(color.getIndex());

        return style;
    }

    public static CellStyle alignMiddle(CellStyle style) {
        alignHorizontally(style, HorizontalAlignment.CENTER);
        alignVertically(style, VerticalAlignment.CENTER);

        return style;
    }

    public static CellStyle alignHorizontally(CellStyle style, HorizontalAlignment alignment) {
        style.setAlignment(alignment);
        return style;
    }

    public static CellStyle alignVertically(CellStyle style, VerticalAlignment alignment) {
        style.setVerticalAlignment(alignment);
        return style;
    }

    public static CellStyle applyBasicHeaderStyle(CellStyle style, Font font) {
        return applyBasicHeaderStyle(style, font, null);
    }

    public static CellStyle applyBasicHeaderStyle(CellStyle style, Font font, String fontName) {
        // Configures a font settings.
        if (fontName != null) font.setFontName(fontName);
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        font.setItalic(false);
        font.setStrikeout(false);

        // Sets up the font
        style.setFont(font);

        drawBorder(style);
        dyeBackground(style);
        alignMiddle(style);

        return style;
    }

    public static CellStyle applyBasicColumnStyle(CellStyle style, Font font) {
        return applyBasicColumnStyle(style, font, null);
    }

    public static CellStyle applyBasicColumnStyle(CellStyle style, Font font, String fontName) {
        // Configures a font settings.
        if (fontName != null) font.setFontName(fontName);
        font.setFontHeightInPoints((short) 10);
        font.setBold(false);
        font.setItalic(false);
        font.setStrikeout(false);

        // Sets up the font
        style.setFont(font);

        drawBorder(style);

        return style;
    }

}
