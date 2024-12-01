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
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Slf4j
@Service
public class FastExcelExportService {

    private static final int CHUNK_SIZE = 1000;

    @Autowired
    private EmployeeRepository employeeRepository;

    @Transactional(readOnly = true)
    public byte[] exportEmployeesToExcelZip() throws IOException {
        ByteArrayOutputStream zipByteStream = new ByteArrayOutputStream();
        try (ZipOutputStream zipOutputStream = new ZipOutputStream(zipByteStream)) {
            
            // Generate timestamp for filename
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String excelFilename = "employees_fastexcel_" + timestamp + ".xlsx";
            
            // Create ZIP entry
            ZipEntry zipEntry = new ZipEntry(excelFilename);
            zipOutputStream.putNextEntry(zipEntry);
            
            // Generate Excel content
            byte[] excelContent = generateExcelContent();
            
            // Write Excel content to ZIP
            zipOutputStream.write(excelContent);
            zipOutputStream.closeEntry();
        }
        
        return zipByteStream.toByteArray();
    }

    private byte[] generateExcelContent() throws IOException {
        List<String> headers = Arrays.asList("ID", "First Name", "Last Name", "Email", "Department", "Salary");
        
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Workbook workbook = new Workbook(outputStream, "Employees", "1.0");
            Worksheet worksheet = workbook.newWorksheet("Employees");

            // Write headers
            for (int i = 0; i < headers.size(); i++) {
                worksheet.value(0, i, headers.get(i));
            }
            
            // Write data using streaming
            AtomicInteger rowNum = new AtomicInteger(1);
            
            try (Stream<Employee> employeeStream = employeeRepository.streamAll()) {
                employeeStream.forEach(employee -> {
                    try {
                        worksheet.value(rowNum.get(), 0, employee.getId());
                        worksheet.value(rowNum.get(), 1, employee.getFirstName());
                        worksheet.value(rowNum.get(), 2, employee.getLastName());
                        worksheet.value(rowNum.get(), 3, employee.getEmail());
                        worksheet.value(rowNum.get(), 4, employee.getDepartment());
                        worksheet.value(rowNum.get(), 5, employee.getSalary());
                        
                        rowNum.incrementAndGet();
                        
                        // Flush every CHUNK_SIZE rows
                        if (rowNum.get() % CHUNK_SIZE == 0) {
                            worksheet.flush();
                        }
                    } catch (IOException e) {
                        throw new RuntimeException("Error writing to Excel file", e);
                    }
                });
            }
            
            // Finish and close the workbook
            workbook.finish();
            return outputStream.toByteArray();
        }
    }
}
