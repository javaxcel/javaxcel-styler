package com.github.javaxcel.styler.role;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

public final class Background {

    public static void drawPattern(CellStyle cellStyle, FillPatternType pattern) {
        cellStyle.setFillPattern(pattern);
    }

    public static void dye(CellStyle cellStyle, IndexedColors color) {
        cellStyle.setFillForegroundColor(color.getIndex());
    }

}
