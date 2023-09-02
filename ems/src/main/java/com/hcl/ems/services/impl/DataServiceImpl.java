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
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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
    @Value("${spring.config.location}")
    private String outputPath;
    private static final String OUTPUT_FILE = "Salary Records.xlsx";

    @Override
    public String loadData(String month, MultipartFile multipartFile)
    {
        try
        {
            File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);
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
    public String deleteColumnsFromMultipleSheets(String startDate, String endDate,
                                                  MultipartFile multipartFile) throws IOException
    {
        List<String> sheetList = getSheetList(startDate, endDate);
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            for(String month:sheetList)
            {
                deleteColumn(month, workbook);
            }
            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + OUTPUT_FILE);
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return "Deleted successfully";
    }

    @Override
    public String removeMergedCells(String startDate, String endDate,
                                    MultipartFile multipartFile) throws IOException
    {
        List<String> sheetList = getSheetList(startDate, endDate);
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            for(String month:sheetList)
            {
                removeMergedCells(month, workbook, 15, 17, 1);
                removeMergedCells(month, workbook, 18, 20, 1);
            }
            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + OUTPUT_FILE);
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return "Unmerged successfully";
    }

    public void deleteColumn(String month, Workbook workbook)
    {
        Sheet sheet = workbook.getSheet(month);
        for(int rowNum = 0; rowNum <= 173; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            for (int columnIndex = 15; columnIndex <= row.getLastCellNum(); columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    row.removeCell(cell);
                }
            }
        }
    }

    public void removeMergedCells(String month, Workbook workbook, int start, int end, int region)
    {
        Sheet sheet = workbook.getSheet(month);
        CellRangeAddress mergedRegion = new CellRangeAddress(0, 0, start, end);

        int firstRow = mergedRegion.getFirstRow();
        int lastRow = mergedRegion.getLastRow();
        int firstCol = mergedRegion.getFirstColumn();
        int lastCol = mergedRegion.getLastColumn();

        sheet.removeMergedRegion(region);

        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
        {
            Row row = sheet.getRow(rowIndex);
            if (row == null)
            {
                row = sheet.createRow(rowIndex);
            }
            for (int colIndex = firstCol; colIndex <= lastCol; colIndex++)
            {
                Cell cell = row.createCell(colIndex);
                cell.setCellValue("");
            }
        }
    }

    @Override
    public String addEps(String startDate, String endDate, MultipartFile multipartFile) throws IOException
    {
        List<String> sheetList = getSheetList(startDate, endDate);
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            for(String month:sheetList)
            {
                addColumn(month, workbook);
            }
            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + OUTPUT_FILE);
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return "EPS added successfully";
    }

    public void addColumn(String month, Workbook workbook)
    {
        Sheet sheet = workbook.getSheet(month);
        sheet.setColumnWidth(15, 7168);

        CellStyle cell1Style = workbook.createCellStyle();
        CellStyle cell2Style = workbook.createCellStyle();
        CellStyle cellStyle3 = workbook.createCellStyle();

        createCellStyles(cell1Style, cell2Style, cellStyle3, workbook);

        Row row1 = sheet.getRow(0);
        Row row2 = sheet.getRow(1);
        formatRows(row1, row2, cell2Style);

        String[] monthArr = month.split("-");
        int mon = Integer.parseInt(monthArr[0]);
        int year = Integer.parseInt(monthArr[1]);

        Cell cell1 = row1.createCell(15);
        Cell cell2 = row2.createCell(15);
        Cell cell3 = row2.createCell(16);
        Cell cell4 = row2.createCell(17);

        formatCells(cell1, cell2, cell3, cell4, cell1Style, cell2Style, List.of(mon,year));

        for(int rowNum = 2; rowNum <= 173; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            Cell totalSalaryCell = row.getCell(14);
            double totalSalary = totalSalaryCell.getNumericCellValue();
            double eps;
            if(year<=2000 || (year == 2001 && mon <= 8))
            {
                if(totalSalary<5000)
                {
                    eps = Math.round(totalSalary*0.0833);
                }
                else
                {
                    eps = Math.round(5000*0.0833);
                }
            }
            else
            {
                if(totalSalary<6500)
                {
                    eps = Math.round(totalSalary*0.0833);
                }
                else
                {
                    eps = Math.round(6500*0.0833);
                }
            }
            Cell cell5 = row.createCell(15);
            cell5.setCellValue(eps);
            cell5.setCellStyle(cellStyle3);

            Cell cell6 = row.createCell(16);
            long epsOnTotalWage = Math.round(totalSalary * 0.0833);
            cell6.setCellValue(epsOnTotalWage);
            cell6.setCellStyle(cellStyle3);

            Cell cell7 = row.createCell(17);
            cell7.setCellValue(epsOnTotalWage - eps);
            cell7.setCellStyle(cellStyle3);
        }
    }

    private void formatCells(Cell cell1, Cell cell2, Cell cell3, Cell cell4,
                             CellStyle cell1Style, CellStyle cell2Style, List<Integer> dateList)
    {
        StringBuilder cell1Value = new StringBuilder("Statutory Ceiling Rs.");
        int year = dateList.get(0);
        int mon = dateList.get(1);
        if(year<=2000 || (year == 2001 && mon <= 8))
        {
            cell1Value.append("5000");
        }
        else
        {
            cell1Value.append("6500");
        }

        cell1.setCellStyle(cell1Style);

        cell2.setCellValue("EPS Deposited");
        cell2.setCellStyle(cell2Style);

        cell3.setCellValue("EPS on Total Wage");
        cell3.setCellStyle(cell2Style);

        cell4.setCellValue("EPS Due");
        cell4.setCellStyle(cell2Style);

        cell1.setCellValue(cell1Value.toString());
    }

    private void formatRows(Row row1, Row row2, CellStyle cell2Style)
    {
        short height = 1024;
        row1.setHeight(height);

        for(int i=0;i<=14;i++)
        {
            Cell cell = row2.getCell(i);
            cell2Style.setWrapText(true);
            cell.setCellStyle(cell2Style);
        }
    }

    private void createCellStyles(CellStyle cell1Style, CellStyle cell2Style,
                                  CellStyle cellStyle3, Workbook workbook)
    {
        String fontStyle = "Calibri";

        cell1Style.setBorderTop(BorderStyle.THIN);
        cell1Style.setBorderRight(BorderStyle.THIN);
        cell1Style.setBorderBottom(BorderStyle.THIN);
        cell1Style.setBorderLeft(BorderStyle.THIN);
        Font font1 = workbook.createFont();
        font1.setFontName(fontStyle);
        font1.setBold(true);
        font1.setFontHeightInPoints((short) 14);
        cell1Style.setFont(font1);
        cell1Style.setAlignment(HorizontalAlignment.CENTER);
        cell1Style.setVerticalAlignment(VerticalAlignment.TOP);

        cell2Style.setBorderTop(BorderStyle.THIN);
        cell2Style.setBorderRight(BorderStyle.THIN);
        cell2Style.setBorderBottom(BorderStyle.THIN);
        cell2Style.setBorderLeft(BorderStyle.THIN);
        Font font2 = workbook.createFont();
        font2.setFontName(fontStyle);
        font2.setBold(true);
        font2.setFontHeightInPoints((short) 12);
        cell2Style.setFont(font2);
        cell2Style.setAlignment(HorizontalAlignment.CENTER);
        cell2Style.setVerticalAlignment(VerticalAlignment.TOP);

        cellStyle3.setBorderTop(BorderStyle.THIN);
        cellStyle3.setBorderRight(BorderStyle.THIN);
        cellStyle3.setBorderBottom(BorderStyle.THIN);
        cellStyle3.setBorderLeft(BorderStyle.THIN);
        Font font3 = workbook.createFont();
        font3.setFontName(fontStyle);
        font3.setFontHeightInPoints((short) 11);
        cellStyle3.setFont(font3);
        cellStyle3.setAlignment(HorizontalAlignment.CENTER);
    }

    @Override
    public List<Epf> verifyData(String month, MultipartFile multipartFile, EpfDto epfDto)
    {
        List<Epf> epfMismatchList = new ArrayList<>();
        try
        {
            File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);
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
