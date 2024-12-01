package com.example.excelexport.controller;

import com.example.excelexport.service.ExcelExportService;
import com.example.excelexport.service.FastExcelExportService;
import com.example.excelexport.service.DirectExcelExportService;
import com.example.excelexport.service.EasyExcelExportService;
import com.example.excelexport.entity.Employee;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelExportController {

    private final ExcelExportService excelExportService;
    private final FastExcelExportService fastExcelExportService;
    private final DirectExcelExportService directExcelExportService;
    private final EasyExcelExportService easyExcelExportService;

    // Export Endpoints
    @GetMapping("/export/zip/poi")
    public ResponseEntity<Resource> exportExcelZipPoi() throws IOException {
        byte[] zipContent = excelExportService.exportEmployeesToExcelZip();
        String filename = "employees_poi_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) + ".zip";
        
        ByteArrayResource resource = new ByteArrayResource(zipContent);

        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
            .contentType(MediaType.parseMediaType("application/zip"))
            .contentLength(zipContent.length)
            .body(resource);
    }

    @GetMapping("/export/zip/fastexcel")
    public ResponseEntity<Resource> exportExcelZipFastExcel() throws IOException {
        byte[] zipContent = fastExcelExportService.exportEmployeesToExcelZip();
        String filename = "employees_fastexcel_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) + ".zip";
        
        ByteArrayResource resource = new ByteArrayResource(zipContent);

        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
            .contentType(MediaType.parseMediaType("application/zip"))
            .contentLength(zipContent.length)
            .body(resource);
    }

    @GetMapping("/export/excel/direct")
    public ResponseEntity<Resource> exportExcelDirect() throws IOException {
        byte[] excelContent = directExcelExportService.exportEmployeesToExcel();
        String filename = "employees_direct_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) + ".xlsx";
        
        ByteArrayResource resource = new ByteArrayResource(excelContent);

        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
            .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
            .contentLength(excelContent.length)
            .body(resource);
    }

    @GetMapping("/export/excel/easyexcel")
    public ResponseEntity<Resource> exportExcelEasyExcel() throws IOException {
        byte[] excelContent = easyExcelExportService.exportEmployeesToExcel();
        String filename = "employees_easyexcel_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) + ".xlsx";
        
        ByteArrayResource resource = new ByteArrayResource(excelContent);

        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
            .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
            .contentLength(excelContent.length)
            .body(resource);
    }

    // Employee Endpoints
    @GetMapping("/employees")
    public List<Employee> getAllEmployees() {
        return excelExportService.getAllEmployees();
    }

    @GetMapping("/employees/{id}")
    public Employee getEmployee(@PathVariable Long id) {
        return excelExportService.getEmployeeById(id);
    }

    @PostMapping("/employees")
    public Employee createEmployee(@RequestBody Employee employee) {
        return excelExportService.saveEmployee(employee);
    }

    @DeleteMapping("/employees/{id}")
    public void deleteEmployee(@PathVariable Long id) {
        excelExportService.deleteEmployee(id);
    }
}
