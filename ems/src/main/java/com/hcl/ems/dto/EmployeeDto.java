package com.hcl.ems.dto;

import lombok.Builder;
import lombok.Data;

import java.util.Map;

@Data
@Builder
public class EmployeeDto
{
    private Integer staffNo;
    private Map<String, Double> salaryMap;
}
