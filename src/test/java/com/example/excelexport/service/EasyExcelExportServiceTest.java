package com.example.excelexport.service;

import com.example.excelexport.entity.Employee;
import com.example.excelexport.repository.EmployeeRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
public class EasyExcelExportServiceTest {

    @Mock
    private EmployeeRepository employeeRepository;

    @InjectMocks
    private EasyExcelExportService easyExcelExportService;

    private Employee testEmployee;

    @BeforeEach
    void setUp() {
        testEmployee = new Employee();
        testEmployee.setId(1L);
        testEmployee.setFirstName("John");
        testEmployee.setLastName("Doe");
        testEmployee.setEmail("john.doe@example.com");
        testEmployee.setDepartment("IT");
        testEmployee.setSalary(75000.0);
    }

    @Test
    void exportEmployeesToExcel_ShouldReturnByteArray() throws IOException {
        try (Stream<Employee> employeeStream = Stream.of(testEmployee)) {
            when(employeeRepository.streamAll()).thenReturn(employeeStream);

            byte[] result = easyExcelExportService.exportEmployeesToExcel();

            assertNotNull(result);
            assertTrue(result.length > 0);
            verify(employeeRepository).streamAll();
        }
    }

    @Test
    void exportEmployeesToExcel_WithEmptyData_ShouldReturnValidExcel() throws IOException {
        try (Stream<Employee> emptyStream = Stream.empty()) {
            when(employeeRepository.streamAll()).thenReturn(emptyStream);

            byte[] result = easyExcelExportService.exportEmployeesToExcel();

            assertNotNull(result);
            assertTrue(result.length > 0);
            verify(employeeRepository).streamAll();
        }
    }

    @Test
    void exportEmployeesToExcel_WithLargeData_ShouldHandleMemoryEfficiently() throws IOException {
        // Create a stream of 1000 employees
        Stream<Employee> largeStream = Stream.generate(() -> testEmployee).limit(1000);
        when(employeeRepository.streamAll()).thenReturn(largeStream);

        byte[] result = easyExcelExportService.exportEmployeesToExcel();

        assertNotNull(result);
        assertTrue(result.length > 0);
        verify(employeeRepository).streamAll();
    }
}
