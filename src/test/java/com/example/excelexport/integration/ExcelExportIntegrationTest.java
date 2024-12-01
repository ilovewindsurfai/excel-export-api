package com.example.excelexport.integration;

import com.example.excelexport.entity.Employee;
import com.example.excelexport.repository.EmployeeRepository;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.web.client.TestRestTemplate;
import org.springframework.boot.test.web.server.LocalServerPort;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.test.context.DynamicPropertyRegistry;
import org.springframework.test.context.DynamicPropertySource;
import org.testcontainers.containers.PostgreSQLContainer;
import org.testcontainers.junit.jupiter.Container;
import org.testcontainers.junit.jupiter.Testcontainers;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest(webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
@Testcontainers
public class ExcelExportIntegrationTest {

    @Container
    static PostgreSQLContainer<?> postgres = new PostgreSQLContainer<>("postgres:14-alpine");

    @LocalServerPort
    private int port;

    @Autowired
    private TestRestTemplate restTemplate;

    @Autowired
    private EmployeeRepository employeeRepository;

    private String baseUrl;

    @DynamicPropertySource
    static void configureProperties(DynamicPropertyRegistry registry) {
        registry.add("spring.datasource.url", postgres::getJdbcUrl);
        registry.add("spring.datasource.username", postgres::getUsername);
        registry.add("spring.datasource.password", postgres::getPassword);
    }

    @BeforeEach
    void setUp() {
        baseUrl = "http://localhost:" + port + "/api/excel";
        employeeRepository.deleteAll();
        createTestData();
    }

    private void createTestData() {
        for (int i = 1; i <= 10; i++) {
            Employee employee = new Employee();
            employee.setFirstName("FirstName" + i);
            employee.setLastName("LastName" + i);
            employee.setEmail("employee" + i + "@example.com");
            employee.setDepartment("Department" + (i % 3 + 1));
            employee.setSalary(50000.0 + (i * 1000));
            employeeRepository.save(employee);
        }
    }

    @Test
    void shouldExportExcelUsingPOI() throws IOException {
        ResponseEntity<byte[]> response = restTemplate.exchange(
                baseUrl + "/export/zip/poi",
                HttpMethod.GET,
                null,
                byte[].class
        );

        assertEquals(HttpStatus.OK, response.getStatusCode());
        assertNotNull(response.getBody());
        
        // Verify ZIP content
        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(response.getBody()))) {
            ZipEntry entry = zis.getNextEntry();
            assertNotNull(entry);
            assertTrue(entry.getName().endsWith(".xlsx"));
            
            // Read Excel content
            Workbook workbook = WorkbookFactory.create(zis);
            assertEquals(10, workbook.getSheetAt(0).getLastRowNum());
        }
    }

    @Test
    void shouldExportExcelUsingEasyExcel() throws IOException {
        ResponseEntity<byte[]> response = restTemplate.exchange(
                baseUrl + "/export/excel/easyexcel",
                HttpMethod.GET,
                null,
                byte[].class
        );

        assertEquals(HttpStatus.OK, response.getStatusCode());
        assertNotNull(response.getBody());
        
        // Verify Excel content
        try (ByteArrayInputStream bis = new ByteArrayInputStream(response.getBody());
             Workbook workbook = WorkbookFactory.create(bis)) {
            assertEquals(10, workbook.getSheetAt(0).getLastRowNum());
        }
    }

    @Test
    void shouldManageEmployeesSuccessfully() {
        // Create new employee
        Employee newEmployee = new Employee();
        newEmployee.setFirstName("John");
        newEmployee.setLastName("Doe");
        newEmployee.setEmail("john.doe@example.com");
        newEmployee.setDepartment("IT");
        newEmployee.setSalary(75000.0);

        ResponseEntity<Employee> createResponse = restTemplate.postForEntity(
                baseUrl + "/employees",
                newEmployee,
                Employee.class
        );

        assertEquals(HttpStatus.OK, createResponse.getStatusCode());
        assertNotNull(createResponse.getBody());
        assertNotNull(createResponse.getBody().getId());

        // Get all employees
        ResponseEntity<List<Employee>> getAllResponse = restTemplate.exchange(
                baseUrl + "/employees",
                HttpMethod.GET,
                null,
                new ParameterizedTypeReference<List<Employee>>() {}
        );

        assertEquals(HttpStatus.OK, getAllResponse.getStatusCode());
        assertNotNull(getAllResponse.getBody());
        assertEquals(11, getAllResponse.getBody().size());

        // Get single employee
        Long employeeId = createResponse.getBody().getId();
        ResponseEntity<Employee> getResponse = restTemplate.getForEntity(
                baseUrl + "/employees/" + employeeId,
                Employee.class
        );

        assertEquals(HttpStatus.OK, getResponse.getStatusCode());
        assertNotNull(getResponse.getBody());
        assertEquals("John", getResponse.getBody().getFirstName());

        // Delete employee
        restTemplate.delete(baseUrl + "/employees/" + employeeId);

        ResponseEntity<List<Employee>> finalGetAllResponse = restTemplate.exchange(
                baseUrl + "/employees",
                HttpMethod.GET,
                null,
                new ParameterizedTypeReference<List<Employee>>() {}
        );

        assertEquals(HttpStatus.OK, finalGetAllResponse.getStatusCode());
        assertEquals(10, finalGetAllResponse.getBody().size());
    }

    @Test
    void shouldHandleLargeDatasetExport() {
        // Create 1000 employees for testing large dataset
        for (int i = 0; i < 1000; i++) {
            Employee employee = new Employee();
            employee.setFirstName("First" + i);
            employee.setLastName("Last" + i);
            employee.setEmail("email" + i + "@example.com");
            employee.setDepartment("Dept" + (i % 10));
            employee.setSalary(50000.0 + i);
            employeeRepository.save(employee);
        }

        // Test POI export
        ResponseEntity<byte[]> poiResponse = restTemplate.exchange(
                baseUrl + "/export/zip/poi",
                HttpMethod.GET,
                null,
                byte[].class
        );
        assertEquals(HttpStatus.OK, poiResponse.getStatusCode());
        assertNotNull(poiResponse.getBody());
        assertTrue(poiResponse.getBody().length > 0);

        // Test EasyExcel export
        ResponseEntity<byte[]> easyExcelResponse = restTemplate.exchange(
                baseUrl + "/export/excel/easyexcel",
                HttpMethod.GET,
                null,
                byte[].class
        );
        assertEquals(HttpStatus.OK, easyExcelResponse.getStatusCode());
        assertNotNull(easyExcelResponse.getBody());
        assertTrue(easyExcelResponse.getBody().length > 0);
    }
}
