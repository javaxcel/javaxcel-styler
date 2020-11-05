package com.github.javaxcel.styler.config;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import static com.github.javaxcel.styler.role.Border.*;

public final class BorderConfigurer {

    private final Configurer configurer;

    private final CellStyle cellStyle;

    BorderConfigurer(Configurer configurer) {
        this.configurer = configurer;
        this.cellStyle = configurer.cellStyle;
    }

    public BorderConfigurer top(BorderStyle border, IndexedColors color) {
        drawTop(cellStyle, border);
        dyeTop(cellStyle, color);
        return this;
    }

    public BorderConfigurer right(BorderStyle border, IndexedColors color) {
        drawRight(cellStyle, border);
        dyeRight(cellStyle, color);
        return this;
    }

    public BorderConfigurer bottom(BorderStyle border, IndexedColors color) {
        drawBottom(cellStyle, border);
        dyeBottom(cellStyle, color);
        return this;
    }

    public BorderConfigurer left(BorderStyle border, IndexedColors color) {
        drawLeft(cellStyle, border);
        dyeLeft(cellStyle, color);
        return this;
    }

    public BorderConfigurer topAndBottom(BorderStyle border, IndexedColors color) {
        drawTop(cellStyle, border);
        dyeTop(cellStyle, color);
        drawBottom(cellStyle, border);
        dyeBottom(cellStyle, color);
        return this;
    }

    public BorderConfigurer leftAndRight(BorderStyle border, IndexedColors color) {
        drawLeft(cellStyle, border);
        dyeLeft(cellStyle, color);
        drawRight(cellStyle, border);
        dyeRight(cellStyle, color);
        return this;
    }

    public BorderConfigurer all(BorderStyle border, IndexedColors color) {
        drawAll(cellStyle, border);
        dyeAll(cellStyle, color);
        return this;
    }

    public Configurer and() {
        return this.configurer;
    }

}
