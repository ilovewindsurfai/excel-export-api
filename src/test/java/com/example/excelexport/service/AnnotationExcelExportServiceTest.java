package com.example.excelexport.service;

import com.example.excelexport.dto.TestDTO;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.when;

@ExtendWith(MockitoExtension.class)
class AnnotationExcelExportServiceTest {

    @Mock
    private MessageSource messageSource;

    @InjectMocks
    private AnnotationExcelExportService service;

    private List<TestDTO> testData;
    private Date testDate;

    @BeforeEach
    void setUp() {
        // Set up test locale
        LocaleContextHolder.setLocale(Locale.ENGLISH);

        // Set up test date
        testDate = new Date();

        // Set up test data
        testData = Arrays.asList(
            TestDTO.builder()
                .id(1L)
                .firstName("John")
                .lastName("Doe")
                .amount(new BigDecimal("1234.56"))
                .registrationDate(testDate)
                .build(),
            TestDTO.builder()
                .id(2L)
                .firstName("Jane")
                .lastName("Smith")
                .amount(new BigDecimal("7890.12"))
                .registrationDate(testDate)
                .build()
        );

        // Set up message source responses
        when(messageSource.getMessage(eq("excel.header.userId"), any(), any(), any())).thenReturn("Test ID");
        when(messageSource.getMessage(eq("excel.header.firstName"), any(), any(), any())).thenReturn("Test First Name");
        when(messageSource.getMessage(eq("excel.header.lastName"), any(), any(), any())).thenReturn("Test Last Name");
        when(messageSource.getMessage(eq("excel.header.amount"), any(), any(), any())).thenReturn("Test Amount");
        when(messageSource.getMessage(eq("excel.header.registrationDate"), any(), any(), any())).thenReturn("Test Date");
        when(messageSource.getMessage(eq("excel.sheet.data"), any(), any(), any())).thenReturn("Test Sheet");
    }

    @Test
    void generateExcelFromDTO_ShouldCreateValidExcelFile() throws IOException {
        // When
        byte[] excelBytes = service.generateExcelFromDTO(testData);

        // Then
        assertNotNull(excelBytes);
        assertTrue(excelBytes.length > 0);

        // Verify Excel content
        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excelBytes))) {
            Sheet sheet = workbook.getSheetAt(0);
            assertNotNull(sheet);

            // Verify headers
            Row headerRow = sheet.getRow(0);
            assertEquals("Test ID", headerRow.getCell(0).getStringCellValue());
            assertEquals("Test First Name", headerRow.getCell(1).getStringCellValue());
            assertEquals("Test Last Name", headerRow.getCell(2).getStringCellValue());
            assertEquals("Test Amount", headerRow.getCell(3).getStringCellValue());
            assertEquals("Test Date", headerRow.getCell(4).getStringCellValue());

            // Verify first data row
            Row firstDataRow = sheet.getRow(1);
            assertEquals(1.0, firstDataRow.getCell(0).getNumericCellValue(), 0.001);
            assertEquals("John", firstDataRow.getCell(1).getStringCellValue());
            assertEquals("Doe", firstDataRow.getCell(2).getStringCellValue());
            assertEquals(1234.56, firstDataRow.getCell(3).getNumericCellValue(), 0.001);
            
            // Verify date formatting
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
            String expectedDate = dateFormat.format(testDate);
            assertEquals(expectedDate, firstDataRow.getCell(4).getStringCellValue());

            // Verify amount formatting
            CellStyle amountStyle = firstDataRow.getCell(3).getCellStyle();
            assertEquals("#,##0.00", amountStyle.getDataFormatString());
        }
    }

    @Test
    void generateExcelFromDTO_WithEmptyList_ShouldThrowException() {
        // When/Then
        assertThrows(IllegalArgumentException.class, () -> 
            service.generateExcelFromDTO(Arrays.asList()));
    }

    @Test
    void generateExcelFromDTO_WithNullList_ShouldThrowException() {
        // When/Then
        assertThrows(IllegalArgumentException.class, () -> 
            service.generateExcelFromDTO(null));
    }

    @Test
    void generateExcelFromDTO_ShouldHandleNullValues() throws IOException {
        // Given
        TestDTO dtoWithNulls = TestDTO.builder()
            .id(1L)
            .firstName(null)
            .lastName("Doe")
            .amount(null)
            .registrationDate(null)
            .build();

        // When
        byte[] excelBytes = service.generateExcelFromDTO(Arrays.asList(dtoWithNulls));

        // Then
        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excelBytes))) {
            Sheet sheet = workbook.getSheetAt(0);
            Row dataRow = sheet.getRow(1);

            assertEquals(1.0, dataRow.getCell(0).getNumericCellValue(), 0.001);
            assertTrue(dataRow.getCell(1) == null || dataRow.getCell(1).getStringCellValue().isEmpty());
            assertEquals("Doe", dataRow.getCell(2).getStringCellValue());
            assertTrue(dataRow.getCell(3) == null || dataRow.getCell(3).getStringCellValue().isEmpty());
            assertTrue(dataRow.getCell(4) == null || dataRow.getCell(4).getStringCellValue().isEmpty());
        }
    }
}
