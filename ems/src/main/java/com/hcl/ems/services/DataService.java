package com.hcl.ems.services;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.hcl.ems.model.Epf;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface DataService
{
    String loadData(String month, MultipartFile multipartFile);
    String loadBulkData(String startDate, String endDate, MultipartFile multipartFile) throws JsonProcessingException;
    List<Epf> verifyData(String month, MultipartFile multipartFile);
    String createTextFiles(MultipartFile multipartFile) throws IOException;
    String createExcelFile(MultipartFile multipartFile) throws IOException;
    String special(String month, MultipartFile multipartFile, MultipartFile multipartFile1);
    String getRatePercentage(String monthYear);
}
