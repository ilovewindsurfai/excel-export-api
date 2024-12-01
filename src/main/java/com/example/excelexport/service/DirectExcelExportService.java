package com.example.excelexport.service;

import com.example.excelexport.entity.Employee;
import com.example.excelexport.repository.EmployeeRepository;
import lombok.extern.slf4j.Slf4j;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

@Slf4j
@Service
public class DirectExcelExportService {

    private static final int CHUNK_SIZE = 1000;

    @Autowired
    private EmployeeRepository employeeRepository;

    @Transactional(readOnly = true)
    public byte[] exportEmployeesToExcel() throws IOException {
        List<String> headers = Arrays.asList("ID", "First Name", "Last Name", "Email", "Department", "Salary");
        
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Workbook workbook = new Workbook(outputStream, "Employees", "1.0");
            Worksheet worksheet = workbook.newWorksheet("Employees");

            // Write headers with styling
            for (int i = 0; i < headers.size(); i++) {
                worksheet.value(0, i, headers.get(i));
                worksheet.style(0, i).bold().fillColor("C0C0C0").set();
                worksheet.width(i, 15); // Set column width
            }
            
            // Write data using streaming
            AtomicInteger rowNum = new AtomicInteger(1);
            
            try (Stream<Employee> employeeStream = employeeRepository.streamAll()) {
                employeeStream.forEach(employee -> {
                    try {
                        // Write data with alternating row colors
                        if (rowNum.get() % 2 == 0) {
                            for (int i = 0; i < 6; i++) {
                                worksheet.style(rowNum.get(), i).fillColor("F5F5F5").set();
                            }
                        }

                        worksheet.value(rowNum.get(), 0, employee.getId());
                        worksheet.value(rowNum.get(), 1, employee.getFirstName());
                        worksheet.value(rowNum.get(), 2, employee.getLastName());
                        worksheet.value(rowNum.get(), 3, employee.getEmail());
                        worksheet.value(rowNum.get(), 4, employee.getDepartment());
                        worksheet.value(rowNum.get(), 5, employee.getSalary());
                        
                        // Add currency format to salary column
                        worksheet.style(rowNum.get(), 5).format("$#,##0.00").set();
                        
                        rowNum.incrementAndGet();
                        
                        // Flush every CHUNK_SIZE rows
                        if (rowNum.get() % CHUNK_SIZE == 0) {
                            worksheet.flush();
                            log.debug("Processed {} rows", rowNum.get());
                        }
                    } catch (IOException e) {
                        throw new RuntimeException("Error writing to Excel file", e);
                    }
                });
            }
            
            // Add auto-filter to headers
            worksheet.range(0, 0, rowNum.get() - 1, headers.size() - 1).autoFilter();
            
            // Freeze the header row
            worksheet.freezePane(1, 0);
            
            // Finish and close the workbook
            workbook.finish();
            return outputStream.toByteArray();
        }
    }
}
