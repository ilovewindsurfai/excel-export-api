package com.example.excelexport.integration;

import com.example.excelexport.entity.Employee;
import com.example.excelexport.repository.EmployeeRepository;
import io.zonky.test.db.AutoConfigureEmbeddedDatabase;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.web.client.TestRestTemplate;
import org.springframework.boot.test.web.server.LocalServerPort;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest(webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
@AutoConfigureEmbeddedDatabase(provider = AutoConfigureEmbeddedDatabase.DatabaseProvider.ZONKY)
public class ExcelExportZonkyIntegrationTest {

    @LocalServerPort
    private int port;

    @Autowired
    private TestRestTemplate restTemplate;

    @Autowired
    private EmployeeRepository employeeRepository;

    private String baseUrl;

    @BeforeEach
    void setUp() {
        baseUrl = "http://localhost:" + port + "/api/excel";
        employeeRepository.deleteAll();
        createTestData();
    }

    private void createTestData() {
        List<Employee> employees = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            Employee employee = new Employee();
            employee.setFirstName("FirstName" + i);
            employee.setLastName("LastName" + i);
            employee.setEmail("email" + i + "@test.com");
            employee.setSalary(50000.0 + i * 1000);
            employees.add(employee);
        }
        employeeRepository.saveAll(employees);
    }

    @Test
    void testExportToExcel_Apache() throws IOException {
        // Test Apache POI export
        ResponseEntity<ByteArrayResource> response = restTemplate.getForEntity(
                baseUrl + "/apache-poi", ByteArrayResource.class);

        assertEquals(HttpStatus.OK, response.getStatusCode());
        assertNotNull(response.getBody());

        // Verify the Excel content
        try (ByteArrayInputStream bis = new ByteArrayInputStream(response.getBody().getByteArray());
             Workbook workbook = WorkbookFactory.create(bis)) {
            assertEquals(1, workbook.getNumberOfSheets());
            assertEquals(101, workbook.getSheetAt(0).getPhysicalNumberOfRows()); // 100 data rows + 1 header
        }
    }

    @Test
    void testExportToExcel_FastExcel() throws IOException {
        // Test FastExcel export
        ResponseEntity<ByteArrayResource> response = restTemplate.getForEntity(
                baseUrl + "/fastexcel", ByteArrayResource.class);

        assertEquals(HttpStatus.OK, response.getStatusCode());
        assertNotNull(response.getBody());

        // Verify the Excel content
        try (ByteArrayInputStream bis = new ByteArrayInputStream(response.getBody().getByteArray());
             Workbook workbook = WorkbookFactory.create(bis)) {
            assertEquals(1, workbook.getNumberOfSheets());
            assertEquals(101, workbook.getSheetAt(0).getPhysicalNumberOfRows());
        }
    }

    @Test
    void testExportToExcel_EasyExcel() throws IOException {
        // Test EasyExcel export
        ResponseEntity<ByteArrayResource> response = restTemplate.getForEntity(
                baseUrl + "/easyexcel", ByteArrayResource.class);

        assertEquals(HttpStatus.OK, response.getStatusCode());
        assertNotNull(response.getBody());

        // Verify the Excel content
        try (ByteArrayInputStream bis = new ByteArrayInputStream(response.getBody().getByteArray());
             Workbook workbook = WorkbookFactory.create(bis)) {
            assertEquals(1, workbook.getNumberOfSheets());
            assertEquals(101, workbook.getSheetAt(0).getPhysicalNumberOfRows());
        }
    }

    @Test
    void testBatchExportToZip() throws IOException {
        // Test batch export to ZIP
        ResponseEntity<ByteArrayResource> response = restTemplate.getForEntity(
                baseUrl + "/batch/10", ByteArrayResource.class);

        assertEquals(HttpStatus.OK, response.getStatusCode());
        assertNotNull(response.getBody());

        // Verify ZIP content
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(response.getBody().getByteArray()))) {
            int fileCount = 0;
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                assertNotNull(entry.getName());
                assertTrue(entry.getName().endsWith(".xlsx"));
                fileCount++;
            }
            assertEquals(10, fileCount); // Verify we have 10 files in the ZIP
        }
    }
}
