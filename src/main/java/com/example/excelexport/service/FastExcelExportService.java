package com.example.excelexport.service;

import com.example.excelexport.entity.Employee;
import com.example.excelexport.repository.EmployeeRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Slf4j
@Service
@RequiredArgsConstructor
public class FastExcelExportService {

    private static final int CHUNK_SIZE = 1000;

    private final EmployeeRepository employeeRepository;
    private final MessageSource messageSource;

    @Transactional(readOnly = true)
    public byte[] exportEmployeesToExcelZip() throws IOException {
        ByteArrayOutputStream zipByteStream = new ByteArrayOutputStream();
        try (ZipOutputStream zipOutputStream = new ZipOutputStream(zipByteStream)) {
            
            // Generate timestamp for filename
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String baseFilename = messageSource.getMessage(
                "excel.filename.users",
                null,
                "employees_fastexcel",
                LocaleContextHolder.getLocale()
            );
            String excelFilename = baseFilename + "_" + timestamp + ".xlsx";
            
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
        Locale currentLocale = LocaleContextHolder.getLocale();
        
        // Get localized headers
        List<String> headers = Arrays.asList(
            messageSource.getMessage("excel.header.userId", null, "ID", currentLocale),
            messageSource.getMessage("excel.header.firstName", null, "First Name", currentLocale),
            messageSource.getMessage("excel.header.lastName", null, "Last Name", currentLocale),
            messageSource.getMessage("excel.header.email", null, "Email", currentLocale),
            messageSource.getMessage("excel.header.department", null, "Department", currentLocale),
            messageSource.getMessage("excel.header.salary", null, "Salary", currentLocale)
        );
        
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            String sheetName = messageSource.getMessage(
                "excel.sheet.users",
                null,
                "Employees",
                currentLocale
            );
            
            Workbook workbook = new Workbook(outputStream, sheetName, "1.0");
            Worksheet worksheet = workbook.newWorksheet(sheetName);

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
