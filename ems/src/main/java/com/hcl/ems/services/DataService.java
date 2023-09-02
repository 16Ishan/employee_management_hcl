package com.hcl.ems.services;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.hcl.ems.dto.EpfDto;
import com.hcl.ems.model.Epf;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface DataService
{
    String loadData(String month, MultipartFile multipartFile);
    String loadBulkData(String month, String endDate, MultipartFile multipartFile) throws JsonProcessingException;
    List<Epf> verifyData(String month, MultipartFile multipartFile, EpfDto epfo);
    String deleteColumnsFromMultipleSheets(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
    String addEps(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
    String removeMergedCells(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
}
