<p align="center" width="40%">
	<img  src="./src/main/resources/main-image.png" alt="Javaxcel Styler">
</p>




<p align="center">
    <img alt="GitHub commit activity" src="https://img.shields.io/github/commit-activity/m/javaxcel/javaxcel-styler">
    <a href="https://lgtm.com/projects/g/javaxcel/javaxcel-styler/context:java"><img alt="Language grade: Java" src="https://img.shields.io/lgtm/grade/java/g/javaxcel/javaxcel-styler.svg?logo=lgtm&logoWidth=18"/></a>
    <a href="https://frontend.code-inspector.com/project/16362/dashboard"><img alt="Code Inspector" src="https://www.code-inspector.com/project/16362/score/svg"></a>
</p>

<p align="center">
    <img alt="GitHub release (latest by date)" src="https://img.shields.io/github/v/release/javaxcel/javaxcel-styler?label=github">
    <img alt="Bintray" src="https://img.shields.io/bintray/v/imsejin/Javaxcel/javaxcel-styler">
    <img alt="Maven Central" src="https://img.shields.io/maven-central/v/com.github.javaxcel/javaxcel-styler">
</p>

<p align="center">
    <img alt="GitHub All Releases" src="https://img.shields.io/github/downloads/javaxcel/javaxcel-styler/total?label=downloads%20at%20github">
    <img alt="Bintray" src="https://img.shields.io/bintray/dt/imsejin/Javaxcel/javaxcel-styler?label=downloads%20at%20bintray">
    <img alt="GitHub" src="https://img.shields.io/github/license/javaxcel/javaxcel-styler">
    <img alt="jdk8" src="https://img.shields.io/badge/jdk-8-orange">
</p>


Javaxcel styler is configurer for decoration of cell style with simple usage.

<br><br>

# Getting started

```xml
<!-- Maven -->
<dependency>
  <groupId>com.github.javaxcel</groupId>
  <artifactId>javaxcel-styler</artifactId>
  <version>${javaxcel.styler.version}</version>
</dependency>
```

```groovy
// Gradle
implementation 'com.github.javaxcel:javaxcel-styler:$javaxcel_styler_version'
```

<br>

<br>

# Examples

```java
public class DefaultHeaderStyleConfig implements ExcelStyleConfig {
    @Override
    public void configure(Configurer configurer) {
        configurer
            	.alignment()
                    .horizontal(HorizontalAlignment.CENTER)
                    .vertical(VerticalAlignment.CENTER)
            	    .and()
                .background(FillPatternType.SOLID_FOREGROUND, IndexedColors.GREY_25_PERCENT)
                .border()
                    .leftAndRight(BorderStyle.THIN, IndexedColors.BLACK)
                    .bottom(BorderStyle.MEDIUM, IndexedColors.BLACK)
                    .and()
                .font()
                    .name("Arial")
                    .size(12)
                    .color(IndexedColors.BLACK)
                    .bold();
    }
}
```

```java
File file = new File("/data", "result.xlsx");

try (FileOutputStream out = new FileOutputStream(file);
        Workbook workbook = new XSSFWorkbook()) {
    Cell cell;
    
    /* ... */
    
    CellStyle headerStyle = workbook.createCellStyle();
    Font font = workbook.createFont();
    
    // Applies custom style configuration to cell style.
    new DefaultHeaderStyleConfig().configure(new Configurer(headerStyle, font));
    cell.setCellStyle(headerStyle);
    
    /* ... */
    
    workbook.write(out);
} catch (IOException e) {
    e.printStackTrace();
}
```

You can configure custom style with implementation or lambda expression.

The implementation of `ExcelStyleConfig` must have default constructor.

