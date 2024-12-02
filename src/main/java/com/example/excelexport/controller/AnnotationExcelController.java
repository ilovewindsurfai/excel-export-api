package com.example.excelexport.controller;

import com.example.excelexport.dto.UserDTO;
import com.example.excelexport.service.AnnotationExcelExportService;
import lombok.RequiredArgsConstructor;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@RestController
@RequestMapping("/api/excel/annotation")
@RequiredArgsConstructor
public class AnnotationExcelController {

    private final AnnotationExcelExportService annotationExcelExportService;
    private final MessageSource messageSource;

    /**
     * Exports a list of DTOs to Excel using annotations
     * @param data List of DTOs to export
     * @param <T> Type of the DTO
     * @return Excel file as byte array
     */
    @PostMapping("/export")
    public <T> ResponseEntity<byte[]> exportToExcel(
            @RequestBody List<T> data,
            @RequestHeader(name = "Accept-Language", required = false) String locale) throws IOException {
            
        byte[] excelContent = annotationExcelExportService.generateExcelFromDTO(data);
        
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String baseFilename = messageSource.getMessage(
            "excel.filename.generic",
            null,
            "data_export",
            LocaleContextHolder.getLocale()
        );
        String filename = baseFilename + "_" + timestamp + ".xlsx";
        
        return createExcelResponse(excelContent, filename);
    }

    /**
     * Example endpoint specifically for UserDTO export
     * @param users List of UserDTOs to export
     * @return Excel file as byte array
     */
    @PostMapping("/users/export")
    public ResponseEntity<byte[]> exportUsers(
            @RequestBody List<UserDTO> users,
            @RequestHeader(name = "Accept-Language", required = false) String locale) throws IOException {
            
        byte[] excelContent = annotationExcelExportService.generateExcelFromDTO(users);
        
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        String baseFilename = messageSource.getMessage(
            "excel.filename.users",
            null,
            "users_export",
            LocaleContextHolder.getLocale()
        );
        String filename = baseFilename + "_" + timestamp + ".xlsx";
        
        return createExcelResponse(excelContent, filename);
    }

    private ResponseEntity<byte[]> createExcelResponse(byte[] content, String filename) {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", filename);
        headers.setContentLength(content.length);
        
        return ResponseEntity.ok()
                .headers(headers)
                .body(content);
    }
}
