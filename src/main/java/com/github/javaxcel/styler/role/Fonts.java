package com.github.javaxcel.styler.role;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

public final class Fonts {

    public static void name(CellStyle cellStyle, Font font, String name) {
        font.setFontName(name);
        cellStyle.setFont(font);
    }

    public static void size(CellStyle cellStyle, Font font, int size) {
        font.setFontHeightInPoints((short) size);
        cellStyle.setFont(font);
    }

    public static void color(CellStyle cellStyle, Font font, IndexedColors color) {
        font.setColor(color.getIndex());
        cellStyle.setFont(font);
    }

    public static void bold(CellStyle cellStyle, Font font) {
        font.setBold(true);
        cellStyle.setFont(font);
    }

    public static void italic(CellStyle cellStyle, Font font) {
        font.setItalic(true);
        cellStyle.setFont(font);
    }

    public static void strikeout(CellStyle cellStyle, Font font) {
        font.setStrikeout(true);
        cellStyle.setFont(font);
    }

    public static void underline(CellStyle cellStyle, Font font, Underline underline) {
        font.setUnderline(underline.value);
        cellStyle.setFont(font);
    }

    public static void offset(CellStyle cellStyle, Font font, Offset offset) {
        font.setTypeOffset(offset.value);
        cellStyle.setFont(font);
    }

    public enum Offset {
        /**
         * No offset.
         */
        NONE(Font.SS_NONE),

        /**
         * Superscript.
         */
        SUPER(Font.SS_SUPER),

        /**
         * Subscript.
         */
        SUB(Font.SS_SUB);

        private final short value;

        Offset(short value) {
            this.value = value;
        }

        public short getValue() {
            return this.value;
        }
    }

    public enum Underline {
        /**
         * No underline.
         */
        NONE(Font.U_NONE),

        /**
         * Single underline.
         */
        SINGLE(Font.U_SINGLE),

        /**
         * Single underline for accounting.
         */
        SINGLE_ACCOUNTING(Font.U_SINGLE_ACCOUNTING),

        /**
         * Double underline.
         */
        DOUBLE(Font.U_DOUBLE),

        /**
         * Double underline for accounting.
         */
        DOUBLE_ACCOUNTING(Font.U_DOUBLE_ACCOUNTING);

        private final byte value;

        Underline(byte value) {
            this.value = value;
        }

        public short getValue() {
            return this.value;
        }
    }

}