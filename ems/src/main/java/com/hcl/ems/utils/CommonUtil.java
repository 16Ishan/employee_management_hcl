package com.hcl.ems.utils;

import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

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
}