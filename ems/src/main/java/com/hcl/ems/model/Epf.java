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
    private Long staffNo;

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

    @Column(name = "cpf")
    private Double cpf;

    @Column(name = "cpf_arrear_ded")
    private Double cpfArrearsDed;

    @Column(name = "total_cpf")
    private Double totalCpf;

    @Column(name = "eps")
    private Double eps;

    @Column(name = "eps_arrear_ded")
    private Double epsArrearsDed;

    @Column(name = "total_eps")
    private Double totalEps;

    @Column(name = "remark")
    private String remark;

    @Column(name = "month_year")
    private String monthYear;
}
