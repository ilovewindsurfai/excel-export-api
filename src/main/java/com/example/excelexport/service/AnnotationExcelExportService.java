package com.example.excelexport.service;

import com.example.excelexport.annotation.ExcelColumn;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

@Slf4j
@Service
@RequiredArgsConstructor
public class AnnotationExcelExportService {
    
    private static final int CHUNK_SIZE = 1000;
    private static final int WINDOW_SIZE = 100;

    private final MessageSource messageSource;

    /**
     * Generates an Excel file from a list of DTOs using ExcelColumn annotations
     */
    public <T> byte[] generateExcelFromDTO(List<T> data) throws IOException {
        if (data == null || data.isEmpty()) {
            throw new IllegalArgumentException("Data cannot be null or empty");
        }

        Class<?> dtoClass = data.get(0).getClass();
        List<Field> annotatedFields = getAnnotatedFields(dtoClass);
        
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(WINDOW_SIZE)) {
            workbook.setCompressTempFiles(true);
            
            String sheetName = messageSource.getMessage(
                "excel.sheet.data", 
                null, 
                "Data", 
                LocaleContextHolder.getLocale()
            );
            SXSSFSheet sheet = workbook.createSheet(sheetName);

            // Create cell styles for different data types
            Map<String, CellStyle> cellStyles = createCellStyles(workbook);

            // Create and populate header row
            createHeaderRow(sheet, annotatedFields);

            // Create data rows
            AtomicInteger rowNum = new AtomicInteger(1);
            for (T item : data) {
                Row row = sheet.createRow(rowNum.getAndIncrement());
                populateRow(row, item, annotatedFields, cellStyles);

                // Flush rows to disk every CHUNK_SIZE rows
                if (rowNum.get() % CHUNK_SIZE == 0) {
                    sheet.flushRows(CHUNK_SIZE);
                }
            }

            // Write to byte array
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        }
    }

    private Map<String, CellStyle> createCellStyles(SXSSFWorkbook workbook) {
        Map<String, CellStyle> styles = new HashMap<>();
        DataFormat dataFormat = workbook.createDataFormat();
        Locale currentLocale = LocaleContextHolder.getLocale();

        // Create amount style
        CellStyle amountStyle = workbook.createCellStyle();
        amountStyle.setDataFormat(dataFormat.getFormat("#,##0.00"));
        styles.put("default_amount", amountStyle);

        return styles;
    }

    private List<Field> getAnnotatedFields(Class<?> dtoClass) {
        List<Field> annotatedFields = new ArrayList<>();
        for (Field field : dtoClass.getDeclaredFields()) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                field.setAccessible(true);
                annotatedFields.add(field);
            }
        }
        
        annotatedFields.sort(Comparator.comparingInt(field -> 
            field.getAnnotation(ExcelColumn.class).order()));
        
        return annotatedFields;
    }

    private void createHeaderRow(SXSSFSheet sheet, List<Field> annotatedFields) {
        Row headerRow = sheet.createRow(0);
        Locale currentLocale = LocaleContextHolder.getLocale();

        for (int i = 0; i < annotatedFields.size(); i++) {
            Cell cell = headerRow.createCell(i);
            ExcelColumn annotation = annotatedFields.get(i).getAnnotation(ExcelColumn.class);
            
            String headerName;
            if (!annotation.messageKey().isEmpty()) {
                headerName = messageSource.getMessage(
                    annotation.messageKey(), 
                    null, 
                    annotation.name().isEmpty() ? annotatedFields.get(i).getName() : annotation.name(),
                    currentLocale
                );
            } else {
                headerName = annotation.name().isEmpty() ? 
                    annotatedFields.get(i).getName() : annotation.name();
            }
            
            cell.setCellValue(headerName);
        }
    }

    private <T> void populateRow(Row row, T item, List<Field> annotatedFields, Map<String, CellStyle> cellStyles) {
        for (int i = 0; i < annotatedFields.size(); i++) {
            Cell cell = row.createCell(i);
            Field field = annotatedFields.get(i);
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            
            try {
                Object value = field.get(item);
                if (value != null) {
                    if (annotation.isAmount()) {
                        setAmountCellValue(cell, value, annotation, cellStyles);
                    } else if (value instanceof Date) {
                        SimpleDateFormat dateFormat = new SimpleDateFormat(
                            annotation.dateFormat(), 
                            LocaleContextHolder.getLocale()
                        );
                        cell.setCellValue(dateFormat.format((Date) value));
                    } else {
                        cell.setCellValue(value.toString());
                    }
                }
            } catch (IllegalAccessException e) {
                log.error("Error accessing field: " + field.getName(), e);
                cell.setCellValue("");
            }
        }
    }

    private void setAmountCellValue(Cell cell, Object value, ExcelColumn annotation, Map<String, CellStyle> cellStyles) {
        if (value instanceof Number) {
            double amount = ((Number) value).doubleValue();
            cell.setCellValue(amount);
            
            // Get or create style for this format
            String styleKey = annotation.amountFormat();
            CellStyle amountStyle = cellStyles.computeIfAbsent(styleKey, k -> {
                CellStyle style = cell.getSheet().getWorkbook().createCellStyle();
                DataFormat dataFormat = cell.getSheet().getWorkbook().createDataFormat();
                style.setDataFormat(dataFormat.getFormat(annotation.amountFormat()));
                return style;
            });
            
            cell.setCellStyle(amountStyle);
        } else {
            cell.setCellValue(value.toString());
        }
    }
}
