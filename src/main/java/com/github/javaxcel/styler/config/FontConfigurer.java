package com.github.javaxcel.styler.config;

import com.github.javaxcel.styler.role.Fonts;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

public final class FontConfigurer {

    private final Configurer configurer;

    private final CellStyle cellStyle;

    private final Font font;

    FontConfigurer(Configurer configurer) {
        this.configurer = configurer;
        this.cellStyle = configurer.cellStyle;
        this.font = configurer.font;
    }

    public FontConfigurer name(String name) {
        Fonts.name(cellStyle, font, name);
        return this;
    }

    public FontConfigurer size(int size) {
        Fonts.size(cellStyle, font, size);
        return this;
    }

    public FontConfigurer color(IndexedColors color) {
        Fonts.color(cellStyle, font, color);
        return this;
    }

    public FontConfigurer bold() {
        Fonts.bold(cellStyle, font);
        return this;
    }

    public FontConfigurer italic() {
        Fonts.italic(cellStyle, font);
        return this;
    }

    public FontConfigurer strikeout() {
        Fonts.strikeout(cellStyle, font);
        return this;
    }

    public FontConfigurer underline(Fonts.Underline underline) {
        Fonts.underline(cellStyle, font, underline);
        return this;
    }

    public FontConfigurer offset(Fonts.Offset offset) {
        Fonts.offset(cellStyle, font, offset);
        return this;
    }

    public Configurer and() {
        return this.configurer;
    }

}
