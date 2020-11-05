package com.github.javaxcel.styler;

import com.github.javaxcel.exception.UnsupportedWorkbookException;
import com.github.javaxcel.out.ExcelWriter;
import com.github.javaxcel.styler.conf.BodyStyleConfig;
import com.github.javaxcel.styler.conf.HeaderStyleConfig;
import com.github.javaxcel.styler.config.Configurer;
import com.github.javaxcel.styler.model.Product;
import com.github.javaxcel.styler.model.factory.MockFactory;
import com.github.javaxcel.util.ExcelUtils;
import io.github.imsejin.common.tool.Stopwatch;
import io.github.imsejin.common.util.DateTimeUtils;
import lombok.Cleanup;
import lombok.SneakyThrows;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class ExcelStylerTest {

    private Stopwatch stopwatch;

    private static long getNumOfModels(Workbook workbook) {
        if (workbook instanceof SXSSFWorkbook) throw new UnsupportedWorkbookException();
        return ExcelUtils.getSheets(workbook).stream().mapToInt(ExcelUtils::getNumOfModels).sum();
    }

    @SneakyThrows
    private static long getNumOfWrittenModels(Class<? extends Workbook> type, File file) {
        @Cleanup
        Workbook workbook = type == HSSFWorkbook.class
                ? new HSSFWorkbook(new FileInputStream(file))
                : new XSSFWorkbook(file);
        return getNumOfModels(workbook);
    }

    @BeforeEach
    public void beforeEach() {
        this.stopwatch = new Stopwatch(TimeUnit.SECONDS);
        stopwatch.start();
    }

    @AfterEach
    public void afterEach() {
        stopwatch.stop();
        System.out.println(stopwatch.getStatistics());
    }

    @Test
    @Order(0)
    @DisplayName("autoResizeColumns")
    @SneakyThrows
    public void autoResizeColumns(@TempDir Path path) {
        // given
        File file = new File(path.toFile(), DateTimeUtils.now() + ".xls");
        @Cleanup
        FileOutputStream out = new FileOutputStream(file);
        @Cleanup
        HSSFWorkbook workbook = new HSSFWorkbook();

        // when
        List<Product> products = MockFactory.generateRandomProducts(1000);
        ExcelWriter.init(workbook, Product.class)
                .adjustSheet((sheet, numOfRows, numOfColumns) -> Utils.autoResizeColumns(sheet, numOfColumns))
                .write(out, products);

        // then
        assertThat(file)
                .exists();
        assertThat(products.size())
                .isEqualTo(getNumOfWrittenModels(HSSFWorkbook.class, file));
    }

    @Test
    @Order(1)
    @DisplayName("applyHeaderStyle")
    @SneakyThrows
    public void applyHeaderStyle(@TempDir Path path) {
        // given
        File file = new File(path.toFile(), DateTimeUtils.now() + ".xlsx");
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
        assertThat(file)
                .exists();
        assertThat(products.size())
                .isEqualTo(getNumOfWrittenModels(XSSFWorkbook.class, file));
    }

    @Test
    @Order(2)
    @DisplayName("applyBodyStyle")
    @SneakyThrows
    public void applyBodyStyle(@TempDir Path path) {
        // given
        File file = new File(path.toFile(), DateTimeUtils.now() + ".xlsx");
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
        assertThat(file)
                .exists();
        assertThat(products.size())
                .isEqualTo(getNumOfWrittenModels(XSSFWorkbook.class, file));
    }

    @Test
    @Order(3)
    @DisplayName("applyMultipleStyle")
    @SneakyThrows
    public void applyMultipleStyle() {
        // given
        File file = new File("/data", "styled.xls");
        @Cleanup
        FileOutputStream out = new FileOutputStream(file);
        @Cleanup
        HSSFWorkbook workbook = new HSSFWorkbook();

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
                .adjustSheet((sheet, numOfRows, numOfColumns) -> {
                    Utils.autoResizeColumns(sheet, numOfColumns);
                    Utils.hideExtraColumns(sheet, numOfColumns);
//                    Utils.hideExtraRows(sheet, numOfRows);
                })
                .write(out, products);

        // then
        assertThat(file)
                .exists();
        assertThat(products.size())
                .isEqualTo(getNumOfWrittenModels(HSSFWorkbook.class, file));
    }

}