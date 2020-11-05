package com.github.javaxcel.styler.config;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import static com.github.javaxcel.styler.role.Backgrounds.drawPattern;
import static com.github.javaxcel.styler.role.Backgrounds.dye;

public class Configurer {

    // Allows detail configurers to access it.
    final CellStyle cellStyle;

    // Allows detail configurers to access it.
    final Font font;

    public Configurer(CellStyle cellStyle, Font font) {
        this.cellStyle = cellStyle;
        this.font = font;
    }

    /**
     * Configures alignments.
     *
     * @return alignment configurer
     */
    public AlignmentConfigurer alignment() {
        return new AlignmentConfigurer(this);
    }

    /**
     * Configures background.
     *
     * @param pattern background pattern
     * @param color   background color
     * @return configurer
     */
    public Configurer background(FillPatternType pattern, IndexedColors color) {
        drawPattern(cellStyle, pattern);
        dye(cellStyle, color);
        return this;
    }

    /**
     * Configures borders.
     *
     * @return border configurer
     */
    public BorderConfigurer border() {
        return new BorderConfigurer(this);
    }

    /**
     * Configures fonts.
     *
     * @return font configurer
     */
    public FontConfigurer font() {
        return new FontConfigurer(this);
    }

}
