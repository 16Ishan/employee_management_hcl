package com.hcl.ems.controller;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.hcl.ems.dto.DataRequest;
import com.hcl.ems.dto.DataRequest2;
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
    @PostMapping(value = "/special", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> special(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(dataService.special(request.getMonth(),
                    request.getFile(), request.getFile1()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @GetMapping(value = "/rate")
    public ResponseEntity<String> getRatePercentage(@RequestParam String monthYear)
    {
        try
        {
            return new ResponseEntity<>(dataService.getRatePercentage(monthYear), HttpStatus.OK);
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

    @PostMapping(value = "/loadBulkDataMulti", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<StringBuilder> loadBulkDataMulti(@ModelAttribute DataRequest2 request)
    {
        StringBuilder response = new StringBuilder();
        try
        {
            request.getFiles().forEach(file -> {
                try {
                    response.append(dataService.loadBulkDataMulti(request.getStartDate(),
                            request.getEndDate(), file))
                            .append("\n");
                } catch (JsonProcessingException e) {
                    throw new RuntimeException(e);
                }
            });
            return new ResponseEntity<>(response, HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(new StringBuilder(e.getMessage()), HttpStatus.INTERNAL_SERVER_ERROR);
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

    @PostMapping(value = "/createTextFiles", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> createTextFiles(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(dataService.createTextFiles(request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            e.printStackTrace();
            return new ResponseEntity<>("", HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/createTextFilesMulti", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> createTextFilesMulti(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(dataService.createTextFilesMulti(request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            e.printStackTrace();
            return new ResponseEntity<>("", HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/createExcelFile", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> createExcelFile(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(dataService.createExcelFile(request.getFile()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            e.printStackTrace();
            return new ResponseEntity<>("", HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}