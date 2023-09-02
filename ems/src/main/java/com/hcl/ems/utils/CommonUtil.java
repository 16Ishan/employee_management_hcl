package com.hcl.ems.utils;

import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CommonUtil
{
    private CommonUtil(){}
    public static File convertMultiPartToFile(MultipartFile file, String folder) throws IOException
    {
        String newFilePath = "";
        String originalFileName = file.getOriginalFilename();

        if (originalFileName != null)
        {
            newFilePath = originalFileName.replace(" ", "_").toLowerCase();
        }

        File convFile = new File(folder + newFilePath);
        try (FileOutputStream fos = new FileOutputStream(convFile))
        {
            fos.write(file.getBytes());
        }
        return convFile;
    }

    public static List<String> getSheetList(String startDate, String endDate)
    {
        List<String> sheetList = new ArrayList<>();
        String[] startDateArr = startDate.split("-");
        String[] endDateArr = endDate.split("-");

        int startMonth = Integer.parseInt(startDateArr[0]);
        int startYear = Integer.parseInt(startDateArr[1]);
        int endMonth = Integer.parseInt(endDateArr[0]);
        int endYear = Integer.parseInt(endDateArr[1]);

        int year = startYear;
        int month = startMonth;
        while(year <= endYear)
        {
            if(month>12)
            {
                month = 1;
                year ++;
            }
            String sheet = (month < 10 ? "0" + month: month) + "-" + year;
            sheetList.add(sheet);
            month++;

            if(month > endMonth && year == endYear)
                break;
        }

        return sheetList;
    }
}