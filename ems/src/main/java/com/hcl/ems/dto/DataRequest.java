package com.hcl.ems.dto;

import lombok.Data;
import org.springframework.web.multipart.MultipartFile;

@Data
public class DataRequest
{
    private String month;
    private String startDate;
    private String endDate;
    private MultipartFile file;
}
