package com.hcl.ems.services;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

public interface SheetService
{
    String deleteColumnsFromMultipleSheets(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
    String addEps(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
    String removeMergedCells(String startDate, String endDate, MultipartFile multipartFile) throws IOException;
}
