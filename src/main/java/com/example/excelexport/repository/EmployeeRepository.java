package com.example.excelexport.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;
import com.example.excelexport.entity.Employee;

import java.util.stream.Stream;

@Repository
public interface EmployeeRepository extends JpaRepository<Employee, Long> {
    // You can add custom query methods here if needed
    // For example:
    // List<Employee> findByDepartment(String department);
    // List<Employee> findBySalaryGreaterThan(Double salary);
    
    @Query("SELECT e FROM Employee e")
    @Transactional(readOnly = true)
    Stream<Employee> streamAll();
}
