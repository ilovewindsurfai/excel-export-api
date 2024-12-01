package com.example.excelexport.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Entity
@Table(name = "employees")
@Data
public class Employee {
    
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @ExcelProperty("ID")
    private Long id;
    
    @ExcelProperty("First Name")
    @ColumnWidth(15)
    private String firstName;
    
    @ExcelProperty("Last Name")
    @ColumnWidth(15)
    private String lastName;
    
    @ExcelProperty("Email")
    @ColumnWidth(25)
    private String email;
    
    @ExcelProperty("Department")
    @ColumnWidth(20)
    private String department;
    
    @ExcelProperty("Salary")
    @ColumnWidth(15)
    private Double salary;
}
