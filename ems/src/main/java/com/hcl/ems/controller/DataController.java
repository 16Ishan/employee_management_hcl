package com.hcl.ems.controller;

import com.hcl.ems.dto.DataRequest;
import com.hcl.ems.dto.EpfDto;
import com.hcl.ems.model.Epf;
import com.hcl.ems.services.DataService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/data")
public class DataController
{
    @Autowired
    private DataService dataService;

    @PostMapping(value = "/load", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> loadData(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(dataService.loadData(request.getMonth(),
                    request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/loadBulkData", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> loadBulkData(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(dataService.loadBulkData(request.getStartDate(),
                    request.getEndDate(), request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @GetMapping(value = "/verify", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<List<Epf>> verifyData(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(dataService.verifyData(request.getMonth(),
                    request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(new ArrayList<>(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}