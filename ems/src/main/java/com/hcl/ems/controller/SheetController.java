package com.hcl.ems.controller;

import com.hcl.ems.dto.DataRequest;
import com.hcl.ems.services.SheetService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/sheet")
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

    @PostMapping(value = "/addMissingSalary", consumes = {MediaType.APPLICATION_JSON_VALUE})
    public ResponseEntity<Map<String, List<String>>> addMissingSalary(@RequestBody DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(sheetService.addMissingSalary(request.getStartDate(),
                    request.getEndDate(), request.getBasicDtoList()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(new HashMap<>(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/addMissingMonths", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> addMissingMonths(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(sheetService.addMissingMonths(request.getStartDate(),
                    request.getEndDate(), request.getFile(), request.getFile1(), request.getFile2()), HttpStatus.OK);
        }
        catch (Exception e)
        {
            e.printStackTrace();
            return new ResponseEntity<>("", HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/missingTotalFormula", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> checkMissingTotalFormula(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(sheetService.missingTotalFormula(request.getStartDate(),
                    request.getEndDate(), request.getFile(), "check"), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping(value = "/addMissingTotalFormula", consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public ResponseEntity<String> addMissingTotalFormula(@ModelAttribute DataRequest request)
    {
        try
        {
            return new ResponseEntity<>(sheetService.missingTotalFormula(request.getStartDate(),
                    request.getEndDate(), request.getFile(), "add"), HttpStatus.OK);
        }
        catch (Exception e)
        {
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}