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
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
public class ExcelExportServiceTest {

    @Mock
    private EmployeeRepository employeeRepository;

    @InjectMocks
    private ExcelExportService excelExportService;

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
    void getAllEmployees_ShouldReturnListOfEmployees() {
        when(employeeRepository.findAll()).thenReturn(Arrays.asList(testEmployee));

        List<Employee> result = excelExportService.getAllEmployees();

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals(testEmployee.getFirstName(), result.get(0).getFirstName());
        verify(employeeRepository).findAll();
    }

    @Test
    void getEmployeeById_ShouldReturnEmployee() {
        when(employeeRepository.findById(1L)).thenReturn(Optional.of(testEmployee));

        Employee result = excelExportService.getEmployeeById(1L);

        assertNotNull(result);
        assertEquals(testEmployee.getFirstName(), result.getFirstName());
        verify(employeeRepository).findById(1L);
    }

    @Test
    void saveEmployee_ShouldReturnSavedEmployee() {
        when(employeeRepository.save(any(Employee.class))).thenReturn(testEmployee);

        Employee result = excelExportService.saveEmployee(testEmployee);

        assertNotNull(result);
        assertEquals(testEmployee.getFirstName(), result.getFirstName());
        verify(employeeRepository).save(testEmployee);
    }

    @Test
    void deleteEmployee_ShouldDeleteEmployee() {
        doNothing().when(employeeRepository).deleteById(1L);

        excelExportService.deleteEmployee(1L);

        verify(employeeRepository).deleteById(1L);
    }

    @Test
    void exportEmployeesToExcelZip_ShouldReturnByteArray() throws IOException {
        try (Stream<Employee> employeeStream = Stream.of(testEmployee)) {
            when(employeeRepository.streamAll()).thenReturn(employeeStream);

            byte[] result = excelExportService.exportEmployeesToExcelZip();

            assertNotNull(result);
            assertTrue(result.length > 0);
            verify(employeeRepository).streamAll();
        }
    }
}
