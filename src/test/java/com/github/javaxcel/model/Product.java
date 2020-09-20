package com.github.javaxcel.model;

import com.github.javaxcel.annotation.ExcelColumn;
import lombok.*;

@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
@EqualsAndHashCode
public class Product {

    @ExcelColumn("상품번호")
    private long serialNumber;

    private String name;

    @ExcelColumn("API_ID")
    private String apiId;

    @ExcelColumn(value = "가로")
    private Double width;

    @ExcelColumn(defaultValue = "(empty)") // Default value is ineffective to primitive type.
    private double depth;

    private double height;

    @ExcelColumn(value = "WEIGHT", defaultValue = "0") // Default value is effective except primitive type.
    private Double weight;

}
