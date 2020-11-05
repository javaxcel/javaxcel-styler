package com.github.javaxcel.styler;

import com.github.javaxcel.styler.conf.BodyStyleConfig;
import com.github.javaxcel.styler.conf.HeaderStyleConfig;
import com.github.javaxcel.styler.config.Configurer;
import com.github.javaxcel.styler.model.Product;
import com.github.javaxcel.styler.model.factory.MockFactory;
import com.github.javaxcel.out.ExcelWriter;
import lombok.Cleanup;
import lombok.SneakyThrows;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

public class ExcelStylerTest {

    @Test
    @SneakyThrows
    public void autoResizeColumns() {
        // given
        File file = new File("/data", "products-styled-0.xlsx");
        @Cleanup
        FileOutputStream out = new FileOutputStream(file);
        @Cleanup
        XSSFWorkbook workbook = new XSSFWorkbook();

        // when
        List<Product> products = MockFactory.generateRandomProducts(1000);
        ExcelWriter.init(workbook, Product.class)
                .adjustSheet((sheet, numOfRows, numOfColumns) -> ExcelStyler.autoResizeColumns(sheet, numOfColumns))
                .write(out, products);

        // then
        assertThat(file).exists();
    }

    @Test
    @SneakyThrows
    public void applyStandardHeaderStyle() {
        // given
        File file = new File("/data", "products-styled-1.xlsx");
        @Cleanup
        FileOutputStream out = new FileOutputStream(file);
        @Cleanup
        XSSFWorkbook workbook = new XSSFWorkbook();

        // when
        List<Product> products = MockFactory.generateRandomProducts(1000);
        ExcelWriter.init(workbook, Product.class)
                .headerStyle((style, font) -> {
                    new HeaderStyleConfig().configure(new Configurer(style, font));
                    return style;
                })
                .write(out, products);

        // then
        assertThat(file).exists();
    }

    @Test
    @SneakyThrows
    public void applyStandardColumnStyle() {
        // given
        File file = new File("/data", "products-styled-2.xlsx");
        @Cleanup
        FileOutputStream out = new FileOutputStream(file);
        @Cleanup
        XSSFWorkbook workbook = new XSSFWorkbook();

        // when
        List<Product> products = MockFactory.generateRandomProducts(1000);
        ExcelWriter.init(workbook, Product.class)
                .columnStyles((style, font) -> {
                    new BodyStyleConfig().configure(new Configurer(style, font));
                    return style;
                })
                .write(out, products);

        // then
        assertThat(file).exists();
    }

    @Test
    @SneakyThrows
    public void multipleStyle() {
        // given
        File file = new File("/data", "products-styled-3.xlsx");
        @Cleanup
        FileOutputStream out = new FileOutputStream(file);
        @Cleanup
        XSSFWorkbook workbook = new XSSFWorkbook();

        // when
        List<Product> products = MockFactory.generateRandomProducts(1000);
        ExcelWriter.init(workbook, Product.class)
                .headerStyle((style, font) -> {
                    new HeaderStyleConfig().configure(new Configurer(style, font));
                    return style;
                })
                .columnStyles((style, font) -> {
                    new BodyStyleConfig().configure(new Configurer(style, font));
                    return style;
                })
                .adjustSheet((sheet, numOfRows, numOfColumns) -> ExcelStyler.autoResizeColumns(sheet, numOfColumns))
                .write(out, products);

        // then
        assertThat(file).exists();
    }

}