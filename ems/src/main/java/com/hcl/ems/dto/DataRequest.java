package com.hcl.ems.dto;

import lombok.Data;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Data
public class DataRequest
{
    private String month;
    private String startDate;
    private String endDate;
    List<BasicDto> basicDtoList;
    private MultipartFile file;
    private MultipartFile file1;
    private MultipartFile file2;
}
