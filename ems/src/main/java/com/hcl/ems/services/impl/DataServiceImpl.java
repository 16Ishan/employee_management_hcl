package com.hcl.ems.services.impl;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.hcl.ems.dto.CellDto;
import com.hcl.ems.dto.EpfDto;
import com.hcl.ems.model.Epf;
import com.hcl.ems.repositories.EpfRepository;
import com.hcl.ems.services.DataService;
import com.hcl.ems.utils.CommonUtil;
import com.hcl.ems.validators.DataValidator;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Collectors;

@Service
@Slf4j
public class DataServiceImpl implements DataService
{
    @Autowired
    private EpfRepository epfRepository;
    @Autowired
    private DataValidator dataValidator;
    @Override
    public String loadData(String month, MultipartFile multipartFile)
    {
        try
        {
            File file = CommonUtil.convertMultiPartToFile(multipartFile, "/app");
            FileInputStream fileInputStream = new FileInputStream(file);
            try(Workbook workbook = WorkbookFactory.create(fileInputStream))
            {
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

                List<Epf> epfList = new ArrayList<>();

                //Read the sheet of the input month e.g. 01-1996
                Sheet sheet = workbook.getSheet(month);
                for(int rowNum = 2; rowNum <= 173; rowNum++)
                {
                    Row row = sheet.getRow(rowNum);
                    Epf epf = new Epf();
                    createEmployee(row, epf, month, evaluator);

                    log.info("Employee added to table: " + epf);
                    epfList.add(epf);
                }
                //Save all employee details read from a sheet
                epfRepository.saveAll(epfList);
            }
            catch (NullPointerException e)
            {
                log.error(e.getMessage());
                e.printStackTrace();
                return "Sheet for the given month: " + month + " not found.";
            }

            return "Data imported successfully";
        }
        catch (IOException e)
        {
            log.error(e.getMessage());
            e.printStackTrace();
            return e.getMessage();
        }
    }

    private List<String> getSheetList(String startDate, String endDate)
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

    public String loadBulkData(String startDate, String endDate, MultipartFile multipartFile) throws JsonProcessingException
    {
        List<String> sheetList = getSheetList(startDate, endDate);
        StringBuilder stringBuilder = new StringBuilder();

        for(String month:sheetList)
        {
            EpfDto epfDto = new EpfDto();
            //Data verification
            List<Epf> epfMismatchList = verifyData(month, multipartFile, epfDto);

            log.info("Data mismatch list: " + epfMismatchList);

            //if the current month's data is not loaded
            if(epfDto.getEpfList().isEmpty())
            {
                stringBuilder.append(month).append(" - ")
                        .append(loadData(month, multipartFile))
                        .append("\n");
            }
            else
            {
                if(epfDto.getEpfList().size() == 172)
                {
                    //if data mismatch happens for current month's data
                    if(!epfMismatchList.isEmpty())
                    {
                        ObjectMapper objectMapper = new ObjectMapper();
                        objectMapper.enable(SerializationFeature.INDENT_OUTPUT);

                        stringBuilder.append("\n")
                                .append(month).append(" - ")
                                .append("Data verification failed" + "\n")
                                .append(objectMapper.writeValueAsString(epfMismatchList))
                                .append("\n\n");
                    }
                    else
                    {
                        stringBuilder.append(month).append(" - ")
                                .append("Data already imported")
                                .append("\n");
                    }
                }
                else
                {
                    stringBuilder.append(month).append(" - ")
                            .append("Data count mismatch")
                            .append("\n");
                }
            }
        }
        return stringBuilder.toString();
    }

    @Override
    public List<Epf> verifyData(String month, MultipartFile multipartFile, EpfDto epfDto)
    {
        List<Epf> epfMismatchList = new ArrayList<>();
        try
        {
            File file = CommonUtil.convertMultiPartToFile(multipartFile, "/app");
            FileInputStream fileInputStream = new FileInputStream(file);
            try(Workbook workbook = WorkbookFactory.create(fileInputStream))
            {
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                Sheet sheet = workbook.getSheet(month);
                List<String> memberIdList = new ArrayList<>();
                for(int rowNum = 2; rowNum <= 173; rowNum++)
                {
                    Row row = sheet.getRow(rowNum);
                    Cell memberIdCell = row.getCell(2);
                    if(memberIdCell != null)
                    {
                        memberIdList.add(memberIdCell.getStringCellValue());
                    }
                }
                List<Epf> epfList = epfRepository.findByMemberIdAndMonthYear(
                        memberIdList, month);
                epfDto.setEpfList(epfList);

                Map<String, Epf> epfMap = epfList.stream()
                        .collect(Collectors.toMap(Epf::getMemberId, epf -> epf));

                for(int rowNum = 2; rowNum <= 173; rowNum++)
                {
                    Row row = sheet.getRow(rowNum);
                    Cell memberIdCell = row.getCell(2);
                    if(memberIdCell != null)
                    {
                        String memberId = memberIdCell.getStringCellValue();
                        if(epfMap.get(memberId) != null)
                            verifyEmployee(row, epfMap.get(memberId), month, evaluator, epfMismatchList);
                    }
                }
            }
            catch (NullPointerException e)
            {
                epfDto.setEpfList(new ArrayList<>());
                log.error(e.getMessage());
                e.printStackTrace();
            }
        }
        catch (IOException e)
        {
            log.error(e.getMessage());
            e.printStackTrace();
        }

        return epfMismatchList;
    }

    private void createEmployee(Row row, Epf epf, String month, FormulaEvaluator evaluator)
    {
        CellDto cellDto = getCellDto(row, evaluator);

        log.info("Row: " + cellDto);

        epf.setSNo(cellDto.getSNo());
        epf.setSlipNo(cellDto.getSlipNo());
        epf.setMemberId(cellDto.getMemberId());
        epf.setName(cellDto.getName());
        epf.setCategory(cellDto.getCategory());
        epf.setStaffNo(cellDto.getStaffNo());
        epf.setBasic(cellDto.getBasic());
        epf.setPp(cellDto.getPp());
        epf.setDaVda(cellDto.getDaVda());
        epf.setBasicArrears(cellDto.getBasicArrears());
        epf.setPpArrears(cellDto.getPpArrears());
        epf.setDaArrears(cellDto.getDaArrears());
        epf.setCpfArrears(cellDto.getCpfArrears());
        epf.setEpsArrears(cellDto.getEpsArrears());
        epf.setTotalSalary(cellDto.getTotalSalary());
        epf.setCpf(cellDto.getCpf());
        epf.setCpfArrearsDed(cellDto.getCpfArrearsDed());
        epf.setTotalCpf(cellDto.getTotalCpf());
        epf.setEps(cellDto.getEps());
        epf.setEpsArrearsDed(cellDto.getEpsArrearsDed());
        epf.setTotalEps(cellDto.getTotalEps());
        epf.setRemark(cellDto.getRemark());
        epf.setMonthYear(month);
    }

    private void verifyEmployee(Row row, Epf epf, String month, FormulaEvaluator evaluator,
                                List<Epf> epfMismatchList)
    {
        CellDto cellDto = getCellDto(row, evaluator);
        Epf epfMismatch = new Epf();
        epfMismatch.setStaffNo(epf.getStaffNo());
        epfMismatch.setName(epf.getName());

        //For any data mismatch that happens
        AtomicBoolean mismatch = new AtomicBoolean(false);

        //Perform verification for all cell data in a row
        dataValidator.verifySnoSlipMemberNameCatStaff(epf, epfMismatch, cellDto, mismatch);
        dataValidator.verifyBasicPpDaBasicArrPpArrDaArr(epf, epfMismatch, cellDto, mismatch);
        dataValidator.verifyCpfArrEpsArrTotalSalCpfCpfArrDedTotalCpf(epf, epfMismatch, cellDto, mismatch);
        dataValidator.verifyEpsEpsArrDedTotalEpsRemarkMonth(epf, epfMismatch, cellDto, mismatch, month);

        if(mismatch.get())
            epfMismatchList.add(epfMismatch);
    }

    private CellDto getCellDto(Row row, FormulaEvaluator evaluator)
    {
        //Read all 22 cells in a row
        Cell sNo = row.getCell(0);
        Cell slipNo = row.getCell(1);
        Cell memberId = row.getCell(2);
        Cell name = row.getCell(3);
        Cell category = row.getCell(4);
        Cell staffNo = row.getCell(5);
        Cell basic = row.getCell(6);
        Cell pp = row.getCell(7);
        Cell daVda = row.getCell(8);
        Cell basicArrears = row.getCell(9);
        Cell ppArrears = row.getCell(10);
        Cell daArrears = row.getCell(11);
        Cell cpfArrears = row.getCell(12);
        Cell epsArrears = row.getCell(13);
        double totalSalary = 0;
        if(evaluator.evaluate(row.getCell(14)) != null)
        {
            totalSalary = evaluator.evaluate(row.getCell(14)).getNumberValue();
        }
        Cell cpf = row.getCell(15);
        Cell cpfArrearsDed = row.getCell(16);
        double totalCpf = 0;
        if(evaluator.evaluate(row.getCell(17)) != null)
        {
            totalCpf = evaluator.evaluate(row.getCell(17)).getNumberValue();
        }
        Cell eps = row.getCell(18);
        Cell epsArrearsDed = row.getCell(19);
        double totalEps = 0;
        if(evaluator.evaluate(row.getCell(20)) != null)
        {
            totalEps = evaluator.evaluate(row.getCell(20)).getNumberValue();
        }
        Cell remark = row.getCell(21);

        return new CellDto((int)sNo.getNumericCellValue(), (int)slipNo.getNumericCellValue(),
                memberId.getStringCellValue(), name.getStringCellValue(), category.getStringCellValue(),
                (long)staffNo.getNumericCellValue(), basic.getNumericCellValue(), pp.getNumericCellValue(),
                daVda.getNumericCellValue(), basicArrears.getNumericCellValue(), ppArrears.getNumericCellValue(),
                daArrears.getNumericCellValue(), cpfArrears.getNumericCellValue(), epsArrears.getNumericCellValue(),
                totalSalary, cpf.getNumericCellValue(), cpfArrearsDed.getNumericCellValue(), totalCpf,
                eps.getNumericCellValue(), epsArrearsDed.getNumericCellValue(), totalEps, remark.getStringCellValue());
    }
}
