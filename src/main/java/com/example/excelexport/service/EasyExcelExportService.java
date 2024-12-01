package com.example.excelexport.service;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.example.excelexport.entity.Employee;
import com.example.excelexport.repository.EmployeeRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

@Slf4j
@Service
public class EasyExcelExportService {

    private static final int BATCH_SIZE = 1000;

    @Autowired
    private EmployeeRepository employeeRepository;

    @Transactional(readOnly = true)
    public byte[] exportEmployeesToExcel() throws IOException {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            // Create header style
            WriteCellStyle headerStyle = new WriteCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
            WriteFont headerFont = new WriteFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 12);
            headerStyle.setWriteFont(headerFont);

            // Create content style
            WriteCellStyle contentStyle = new WriteCellStyle();
            contentStyle.setWrapped(true);

            // Combine the styles
            HorizontalCellStyleStrategy styleStrategy = new HorizontalCellStyleStrategy(headerStyle, contentStyle);

            // Configure EasyExcel
            EasyExcel.write(outputStream, Employee.class)
                    .registerWriteHandler(styleStrategy)
                    .sheet("Employees")
                    .useDefaultStyle(false)
                    .doWrite(this::fetchData);

            return outputStream.toByteArray();
        }
    }

    private Stream<Employee> fetchData() {
        return employeeRepository.streamAll();
    }
}
