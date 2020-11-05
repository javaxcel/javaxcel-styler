package com.github.javaxcel.styler.role;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public final class Alignment {

    public static void horizontal(CellStyle cellStyle, HorizontalAlignment alignment) {
        cellStyle.setAlignment(alignment);
    }

    public static void vertical(CellStyle cellStyle, VerticalAlignment alignment) {
        cellStyle.setVerticalAlignment(alignment);
    }

}
