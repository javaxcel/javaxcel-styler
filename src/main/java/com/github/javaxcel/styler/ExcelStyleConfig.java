package com.github.javaxcel.styler;

import com.github.javaxcel.styler.config.Configurer;

@FunctionalInterface
public interface ExcelStyleConfig {

    void configure(Configurer configurer);

}
