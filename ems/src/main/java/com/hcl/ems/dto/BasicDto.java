package com.hcl.ems.dto;

import lombok.Data;

import java.util.List;

@Data
public class BasicDto
{
    private String memberId;
    private List<Double> basicList;
}
