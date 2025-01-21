package com.hcl.ems.model;

import com.fasterxml.jackson.annotation.JsonInclude;
import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Entity
@Table(name = "epf", schema = "hcl")
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
@Builder
public class Epf
{
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "id")
    private Integer id;

    @Column(name = "sno")
    private Integer sNo;

    @Column(name = "slip_no")
    private Integer slipNo;

    @Column(name = "member_id")
    private String memberId;

    @Column(name = "name")
    private String name;

    @Column(name = "category")
    private String category;

    @Column(name = "staff_no")
    private String staffNo;

    @Column(name = "basic")
    private Double basic;

    @Column(name = "pp")
    private Double pp;

    @Column(name = "da_vda")
    private Double daVda;

    @Column(name = "basic_arrear")
    private Double basicArrears;

    @Column(name = "pp_arrear")
    private Double ppArrears;

    @Column(name = "da_arrear")
    private Double daArrears;

    @Column(name = "cpf_arrear")
    private Double cpfArrears;

    @Column(name = "eps_arrear")
    private Double epsArrears;

    @Column(name = "total_salary")
    private Double totalSalary;

    @Column(name = "month_year")
    private String monthYear;

    @Column(name = "eps_deposited")
    private Long epsDeposited;

    @Column(name = "eps_on_total_wage")
    private Long epsOnTotalWage;

    @Column(name = "additional_contribution")
    private Long additionalContribution;

    @Column(name = "eps_due")
    private Long epsDue;

    @Column(name = "rate")
    private double rate;

    @Column(name = "final_due")
    private Long finalDue;
}
