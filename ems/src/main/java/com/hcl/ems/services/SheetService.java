package com.hcl.ems.services;

import com.hcl.ems.dto.BasicDto;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface SheetService
{
    String deleteColumnsFromMultipleSheets(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
    String addEps(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
    Map<String, List<String>> addMissingSalary(String startDate, String endDate, List<BasicDto> basicDtoList) throws IOException;
    String addMissingMonths(String startDate, String endDate, MultipartFile file1,
                            MultipartFile file2, MultipartFile file3) throws IOException;
    String removeMergedCells(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
    String missingTotalFormula(String startDate, String endDate, MultipartFile multipartFile, String action) throws IOException;
}
