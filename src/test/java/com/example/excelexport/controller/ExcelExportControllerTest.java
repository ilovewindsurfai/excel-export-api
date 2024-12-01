package com.example.excelexport.controller;

import com.example.excelexport.entity.Employee;
import com.example.excelexport.service.DirectExcelExportService;
import com.example.excelexport.service.EasyExcelExportService;
import com.example.excelexport.service.ExcelExportService;
import com.example.excelexport.service.FastExcelExportService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;

import java.util.Arrays;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

@WebMvcTest(ExcelExportController.class)
public class ExcelExportControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @MockBean
    private ExcelExportService excelExportService;

    @MockBean
    private FastExcelExportService fastExcelExportService;

    @MockBean
    private DirectExcelExportService directExcelExportService;

    @MockBean
    private EasyExcelExportService easyExcelExportService;

    private Employee testEmployee;
    private byte[] testExcelContent;

    @BeforeEach
    void setUp() {
        testEmployee = new Employee();
        testEmployee.setId(1L);
        testEmployee.setFirstName("John");
        testEmployee.setLastName("Doe");
        testEmployee.setEmail("john.doe@example.com");
        testEmployee.setDepartment("IT");
        testEmployee.setSalary(75000.0);

        testExcelContent = "Test Excel Content".getBytes();
    }

    @Test
    void exportExcelZipPoi_ShouldReturnZipFile() throws Exception {
        when(excelExportService.exportEmployeesToExcelZip()).thenReturn(testExcelContent);

        mockMvc.perform(get("/api/excel/export/zip/poi"))
                .andExpect(status().isOk())
                .andExpect(content().contentType("application/zip"))
                .andExpect(header().exists("Content-Disposition"))
                .andExpect(content().bytes(testExcelContent));
    }

    @Test
    void exportExcelZipFastExcel_ShouldReturnZipFile() throws Exception {
        when(fastExcelExportService.exportEmployeesToExcelZip()).thenReturn(testExcelContent);

        mockMvc.perform(get("/api/excel/export/zip/fastexcel"))
                .andExpect(status().isOk())
                .andExpect(content().contentType("application/zip"))
                .andExpect(header().exists("Content-Disposition"))
                .andExpect(content().bytes(testExcelContent));
    }

    @Test
    void exportExcelDirect_ShouldReturnExcelFile() throws Exception {
        when(directExcelExportService.exportEmployeesToExcel()).thenReturn(testExcelContent);

        mockMvc.perform(get("/api/excel/export/excel/direct"))
                .andExpect(status().isOk())
                .andExpect(content().contentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .andExpect(header().exists("Content-Disposition"))
                .andExpect(content().bytes(testExcelContent));
    }

    @Test
    void exportExcelEasyExcel_ShouldReturnExcelFile() throws Exception {
        when(easyExcelExportService.exportEmployeesToExcel()).thenReturn(testExcelContent);

        mockMvc.perform(get("/api/excel/export/excel/easyexcel"))
                .andExpect(status().isOk())
                .andExpect(content().contentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .andExpect(header().exists("Content-Disposition"))
                .andExpect(content().bytes(testExcelContent));
    }

    @Test
    void getAllEmployees_ShouldReturnEmployeeList() throws Exception {
        when(excelExportService.getAllEmployees()).thenReturn(Arrays.asList(testEmployee));

        mockMvc.perform(get("/api/excel/employees"))
                .andExpect(status().isOk())
                .andExpect(content().contentType(MediaType.APPLICATION_JSON))
                .andExpect(jsonPath("$[0].id").value(testEmployee.getId()))
                .andExpect(jsonPath("$[0].firstName").value(testEmployee.getFirstName()))
                .andExpect(jsonPath("$[0].lastName").value(testEmployee.getLastName()));
    }

    @Test
    void getEmployee_ShouldReturnEmployee() throws Exception {
        when(excelExportService.getEmployeeById(1L)).thenReturn(testEmployee);

        mockMvc.perform(get("/api/excel/employees/1"))
                .andExpect(status().isOk())
                .andExpect(content().contentType(MediaType.APPLICATION_JSON))
                .andExpect(jsonPath("$.id").value(testEmployee.getId()))
                .andExpect(jsonPath("$.firstName").value(testEmployee.getFirstName()))
                .andExpect(jsonPath("$.lastName").value(testEmployee.getLastName()));
    }

    @Test
    void createEmployee_ShouldReturnCreatedEmployee() throws Exception {
        when(excelExportService.saveEmployee(any(Employee.class))).thenReturn(testEmployee);

        mockMvc.perform(post("/api/excel/employees")
                .contentType(MediaType.APPLICATION_JSON)
                .content("{\"firstName\":\"John\",\"lastName\":\"Doe\",\"email\":\"john.doe@example.com\",\"department\":\"IT\",\"salary\":75000.0}"))
                .andExpect(status().isOk())
                .andExpect(content().contentType(MediaType.APPLICATION_JSON))
                .andExpect(jsonPath("$.firstName").value(testEmployee.getFirstName()))
                .andExpect(jsonPath("$.lastName").value(testEmployee.getLastName()));
    }

    @Test
    void deleteEmployee_ShouldReturnNoContent() throws Exception {
        mockMvc.perform(delete("/api/excel/employees/1"))
                .andExpect(status().isOk());
    }
}
