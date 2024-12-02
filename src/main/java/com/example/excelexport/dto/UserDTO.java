package com.example.excelexport.dto;

import com.example.excelexport.annotation.ExcelColumn;
import lombok.Data;
import java.util.Date;

@Data
public class UserDTO {
    @ExcelColumn(messageKey = "excel.header.userId", order = 1)
    private Long id;

    @ExcelColumn(messageKey = "excel.header.firstName", order = 2)
    private String firstName;

    @ExcelColumn(messageKey = "excel.header.lastName", order = 3)
    private String lastName;

    @ExcelColumn(messageKey = "excel.header.email", order = 4)
    private String email;

    @ExcelColumn(messageKey = "excel.header.registrationDate", order = 5, dateFormat = "dd/MM/yyyy")
    private Date registrationDate;

    @ExcelColumn(messageKey = "excel.header.active", order = 6)
    private boolean active;
}
