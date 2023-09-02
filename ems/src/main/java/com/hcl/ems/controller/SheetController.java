package com.hcl.ems.controller;

import com.hcl.ems.dto.DataRequest;
import com.hcl.ems.services.SheetService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/data")
public class SheetController
{
    @Autowired
    private SheetService sheetService;

    @PostMapping(value = "/deleteColumns", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> deleteColumns(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(sheetService.deleteColumnsFromMultipleSheets(request.getStartDate(),
                    request.getEndDate(), request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/removeMergedCells", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> removeMergedCells(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(sheetService.removeMergedCells(request.getStartDate(),
                    request.getEndDate(), request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/addEps", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> addEps(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(sheetService.addEps(request.getStartDate(),
                    request.getEndDate(), request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}