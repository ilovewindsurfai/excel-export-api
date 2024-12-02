package com.example.excelexport.dto;

import com.example.excelexport.annotation.ExcelColumn;
import lombok.Builder;
import lombok.Data;
import java.math.BigDecimal;
import java.util.Date;

@Data
@Builder
public class TestDTO {
    @ExcelColumn(messageKey = "excel.header.userId", order = 1)
    private Long id;

    @ExcelColumn(messageKey = "excel.header.firstName", order = 2)
    private String firstName;

    @ExcelColumn(messageKey = "excel.header.lastName", order = 3)
    private String lastName;

    @ExcelColumn(messageKey = "excel.header.amount", order = 4, isAmount = true, amountFormat = "#,##0.00")
    private BigDecimal amount;

    @ExcelColumn(messageKey = "excel.header.registrationDate", order = 5, dateFormat = "dd/MM/yyyy")
    private Date registrationDate;
}
