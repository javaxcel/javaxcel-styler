package com.github.javaxcel.styler.config;

import com.github.javaxcel.styler.role.Alignments;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public final class AlignmentConfigurer {

    private final Configurer configurer;

    private final CellStyle cellStyle;

    AlignmentConfigurer(Configurer configurer) {
        this.configurer = configurer;
        this.cellStyle = configurer.cellStyle;
    }

    public AlignmentConfigurer horizontal(HorizontalAlignment horizontal) {
        Alignments.horizontal(this.cellStyle, horizontal);
        return this;
    }

    public AlignmentConfigurer vertical(VerticalAlignment vertical) {
        Alignments.vertical(this.cellStyle, vertical);
        return this;
    }

    public Configurer and() {
        return this.configurer;
    }

}
