package com.github.javaxcel.styler.conf;

import com.github.javaxcel.styler.config.Configurer;
import com.github.javaxcel.styler.ExcelStyleConfig;
import com.github.javaxcel.styler.role.Fonts;
import org.apache.poi.ss.usermodel.*;

public class CustomExcelStyleConfig implements ExcelStyleConfig {

    @Override
    public void configure(Configurer configurer) {
        configurer.alignment()
                .horizontal(HorizontalAlignment.CENTER)
                .vertical(VerticalAlignment.CENTER)
                .and()
                .background(FillPatternType.SOLID_FOREGROUND, IndexedColors.GREY_25_PERCENT)
                .border()
                .top(BorderStyle.THIN, IndexedColors.BLACK)
                .right(BorderStyle.THIN, IndexedColors.BLACK)
                .bottom(BorderStyle.THIN, IndexedColors.BLACK)
                .left(BorderStyle.THIN, IndexedColors.BLACK)
                .topAndBottom(BorderStyle.THIN, IndexedColors.BLACK)
                .leftAndRight(BorderStyle.THIN, IndexedColors.BLACK)
                .all(BorderStyle.THIN, IndexedColors.BLACK)
                .and()
                .font()
                .name("NanumGothic")
                .size(12)
                .color(IndexedColors.BLACK)
                .bold()
                .italic()
                .strikeout()
                .underline(Fonts.Underline.DOUBLE_ACCOUNTING)
                .offset(Fonts.Offset.SUPER);
    }

}
