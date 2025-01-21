package com.hcl.ems.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@AllArgsConstructor
@Builder
public class CellDto
{
    private Integer sNo;
    private Integer slipNo;
    private String memberId;
    private String name;
    private String category;
    private String staffNo;
    private Double basic;
    private Double pp;
    private Double daVda;
    private Double basicArrears;
    private Double ppArrears;
    private Double daArrears;
    private Double cpfArrears;
    private Double epsArrears;
    private Double totalSalary;
    private Long epsDeposited;
    private Long epsOnTotalWage;
    private Long additionalContribution;
    private Long epsDue;
    private Double rate;
}
