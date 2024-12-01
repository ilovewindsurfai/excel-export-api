package com.example.excelexport.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import com.example.excelexport.repository.EmployeeRepository;
import com.example.excelexport.entity.Employee;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Arrays;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Slf4j
@Service
public class ExcelExportService {
    
    private static final int CHUNK_SIZE = 1000;
    private static final int WINDOW_SIZE = 100;

    @Autowired
    private EmployeeRepository employeeRepository;

    public List<Employee> getAllEmployees() {
        return employeeRepository.findAll();
    }

    public Employee getEmployeeById(Long id) {
        return employeeRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Employee not found with id: " + id));
    }

    public Employee saveEmployee(Employee employee) {
        return employeeRepository.save(employee);
    }

    public void deleteEmployee(Long id) {
        employeeRepository.deleteById(id);
    }

    public byte[] generateExcel(Stream<List<String>> dataStream, List<String> headers) throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(WINDOW_SIZE)) {
            workbook.setCompressTempFiles(true);
            SXSSFSheet sheet = workbook.createSheet("Data");
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
            }

            AtomicInteger rowNum = new AtomicInteger(1);
            
            // Process data in chunks
            dataStream.forEach(rowData -> {
                Row row = sheet.createRow(rowNum.getAndIncrement());
                for (int i = 0; i < rowData.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(rowData.get(i));
                }
                
                // Flush rows to disk every CHUNK_SIZE rows
                if (rowNum.get() % CHUNK_SIZE == 0) {
                    try {
                        sheet.flushRows(CHUNK_SIZE);
                    } catch (IOException e) {
                        log.error("Error flushing rows to disk", e);
                        throw new RuntimeException("Error while generating Excel file", e);
                    }
                }
            });

            // Write to byte array
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        }
    }

    @Transactional(readOnly = true)
    public byte[] exportEmployeesToExcel() throws IOException {
        List<String> headers = Arrays.asList("ID", "First Name", "Last Name", "Email", "Department", "Salary");
        
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(WINDOW_SIZE)) {
            workbook.setCompressTempFiles(true);
            SXSSFSheet sheet = workbook.createSheet("Employees");
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
            }

            // Create data rows using streaming
            AtomicInteger rowNum = new AtomicInteger(1);
            
            try (Stream<Employee> employeeStream = employeeRepository.streamAll()) {
                employeeStream.forEach(employee -> {
                    Row row = sheet.createRow(rowNum.getAndIncrement());
                    row.createCell(0).setCellValue(employee.getId());
                    row.createCell(1).setCellValue(employee.getFirstName());
                    row.createCell(2).setCellValue(employee.getLastName());
                    row.createCell(3).setCellValue(employee.getEmail());
                    row.createCell(4).setCellValue(employee.getDepartment());
                    row.createCell(5).setCellValue(employee.getSalary());
                    
                    // Flush rows to disk every CHUNK_SIZE rows
                    if (rowNum.get() % CHUNK_SIZE == 0) {
                        try {
                            sheet.flushRows(CHUNK_SIZE);
                        } catch (IOException e) {
                            throw new RuntimeException("Error flushing rows to disk", e);
                        }
                    }
                });
            }
            
            // Write to byte array
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        }
    }

    @Transactional(readOnly = true)
    public byte[] exportEmployeesToExcelZip() throws IOException {
        ByteArrayOutputStream zipByteStream = new ByteArrayOutputStream();
        try (ZipOutputStream zipOutputStream = new ZipOutputStream(zipByteStream)) {
            
            // Generate timestamp for filename
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String excelFilename = "employees_" + timestamp + ".xlsx";
            
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
        
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(WINDOW_SIZE)) {
            workbook.setCompressTempFiles(true);
            SXSSFSheet sheet = workbook.createSheet("Employees");
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
            }

            // Create data rows using streaming
            AtomicInteger rowNum = new AtomicInteger(1);
            
            try (Stream<Employee> employeeStream = employeeRepository.streamAll()) {
                employeeStream.forEach(employee -> {
                    Row row = sheet.createRow(rowNum.getAndIncrement());
                    row.createCell(0).setCellValue(employee.getId());
                    row.createCell(1).setCellValue(employee.getFirstName());
                    row.createCell(2).setCellValue(employee.getLastName());
                    row.createCell(3).setCellValue(employee.getEmail());
                    row.createCell(4).setCellValue(employee.getDepartment());
                    row.createCell(5).setCellValue(employee.getSalary());
                    
                    // Flush rows to disk every CHUNK_SIZE rows
                    if (rowNum.get() % CHUNK_SIZE == 0) {
                        try {
                            sheet.flushRows(CHUNK_SIZE);
                        } catch (IOException e) {
                            throw new RuntimeException("Error flushing rows to disk", e);
                        }
                    }
                });
            }
            
            // Write to byte array
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        }
    }
}
