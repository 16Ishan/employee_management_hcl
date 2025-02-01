package com.hcl.ems.services.impl;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.hcl.ems.dto.CellDto;
import com.hcl.ems.dto.EpfDto;
import com.hcl.ems.dto.Staff;
import com.hcl.ems.model.Epf;
import com.hcl.ems.repositories.EpfCustomRepository;
import com.hcl.ems.repositories.EpfRepository;
import com.hcl.ems.services.DataService;
import com.hcl.ems.services.views.IEpfView;
import com.hcl.ems.utils.CommonUtil;
import com.hcl.ems.validators.DataValidator;
import lombok.Builder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.Month;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Collectors;

import static com.hcl.ems.services.impl.SheetServiceImpl.createCellStyle;
import static com.hcl.ems.services.impl.SheetServiceImpl.createHeaderCellStyle;

@Service
@Slf4j
public class DataServiceImpl implements DataService
{
    @Autowired
    private EpfRepository epfRepository;
    @Autowired
    private EpfCustomRepository epfCustomRepository;
    @Autowired
    private DataValidator dataValidator;
    @Value("${spring.config.location}")
    private String outputPath;

    private static final String SEPARATOR = "#~#";

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

                Sheet sheet = workbook.getSheet(month);
                for(int rowNum = 2; rowNum <= 180; rowNum++)
                {
                    Row row = sheet.getRow(rowNum);
                    Epf epf = createEmployee(row, month, evaluator);

                    log.info("Employee added to table: " + epf);
                    epfList.add(epf);
                    log.info(rowNum+"");
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

    public String loadDataMulti(MultipartFile multipartFile)
    {
        try
        {
            File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);
            FileInputStream fileInputStream = new FileInputStream(file);
            try(Workbook workbook = WorkbookFactory.create(fileInputStream))
            {
                List<Epf> epfList = new ArrayList<>();

                Sheet sheet = workbook.getSheetAt(0);
                Row memberIdRow = sheet.getRow(3);
                Row staffRow = sheet.getRow(2);
                Cell staffCell = staffRow.getCell(6);
                String staffNumber = String.valueOf((int) staffCell.getNumericCellValue());
                Cell memberIdCell1 = memberIdRow.getCell(6);
                Cell memberIdCell2 = memberIdRow.getCell(7);

                String memberIdStr1 = memberIdCell1.getStringCellValue();
                String memberIdStr2 = String.valueOf((int)memberIdCell2.getNumericCellValue());

                String memberId = memberIdStr1.concat(memberIdStr2);
                for(int rowNum = 5; rowNum < 260; rowNum++)
                {
                    Row row = sheet.getRow(rowNum);
                    Epf epf = createEmployeeMulti(row, memberId, staffNumber);

                    log.info("Employee added to table: " + epf);
                    epfList.add(epf);
                    log.info(rowNum+"");
                }
                //Save all employee details read from a sheet
                epfRepository.saveAll(epfList);
            }
            catch (NullPointerException e)
            {
                log.error(e.getMessage());
                e.printStackTrace();
                return "Sheet for the given month:  not found.";
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

    public String special(String month, MultipartFile multipartFile, MultipartFile multipartFile1)
    {
        try
        {
            File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);
            File file1 = CommonUtil.convertMultiPartToFile(multipartFile1, outputPath);
            FileInputStream fileInputStream = new FileInputStream(file);
            FileInputStream fileInputStream1 = new FileInputStream(file1);

            List<String> memberIdList = new ArrayList<>();
            try(Workbook workbook1 = WorkbookFactory.create(fileInputStream1))
            {
                Sheet sheet1 = workbook1.getSheet("Sheet1");
                for(int rowNum = 1; rowNum <= 176; rowNum++)
                {
                    Row row = sheet1.getRow(rowNum);
                    Cell cell = row.getCell(1);
                    String memberId = cell.getStringCellValue();
                    memberId = memberId.substring(18);
                    if(memberId.charAt(0) == '0')
                        memberId = memberId.substring(1);
                    memberIdList.add(memberId);
                }
            }
            try(Workbook workbook = WorkbookFactory.create(fileInputStream))
            {
                Sheet sheet = workbook.getSheet("Sheet1");
                for(int rowNum = 715; rowNum >= 1; rowNum--)
                {
                    Row row = sheet.getRow(rowNum);
                    Cell cell = row.getCell(3);
                    String memberId = String.valueOf(cell.getNumericCellValue());
                    memberId = memberId.substring(0,memberId.lastIndexOf('.'));
                    if(!memberIdList.contains(memberId))
                    {
                        sheet.removeRow(row);
                        // Shift remaining rows up to fill the gap
                        sheet.shiftRows(rowNum + 1, sheet.getLastRowNum(), -1);
                    }
                }

                // Write the changes back to the Excel file
                try (FileOutputStream fos = new FileOutputStream("D:/Downloads/Active Filtered.xlsx")) {
                    workbook.write(fos);
                }
            }
            catch (NullPointerException e)
            {
                log.error(e.getMessage());
                return "Sheet for the given month: " + month + " not found.";
            }

            return "Data filtered successfully";
        }
        catch (IOException e)
        {
            log.error(e.getMessage());
            return e.getMessage();
        }
    }
    @Override
    public String loadBulkDataMulti(String startDate, String endDate, MultipartFile multipartFile) throws JsonProcessingException
    {
        StringBuilder stringBuilder = new StringBuilder();
        loadDataMulti(multipartFile);
        stringBuilder.append("Done");
        return stringBuilder.toString();
    }


    @Override
    public String loadBulkData(String startDate, String endDate, MultipartFile multipartFile) throws JsonProcessingException
    {
        List<String> sheetList = CommonUtil.getSheetList(startDate, endDate);
        StringBuilder stringBuilder = new StringBuilder();

        for(String month:sheetList)
        {
            EpfDto epfDto = new EpfDto();
            stringBuilder.append(month).append(" - ")
                    .append(loadData(month, multipartFile))
                    .append("\n");
            //Data verification
            /*List<Epf> epfMismatchList = verifyData(month, multipartFile, epfDto);

            //log.info("Data mismatch list: " + epfMismatchList);

            //if the current month's data is not loaded
            if(epfDto.getEpfList().isEmpty())
            {
                stringBuilder.append(month).append(" - ")
                        .append(loadData(month, multipartFile))
                        .append("\n");
            }
            else
            {
                if(epfDto.getEpfList().size() == 179)
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
            }*/
        }
        return stringBuilder.toString();
    }

    public List<Epf> verifyData(String month, MultipartFile multipartFile)
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
                for(int rowNum = 2; rowNum <= 180; rowNum++)
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

                Map<String, Epf> epfMap = epfList.stream()
                        .collect(Collectors.toMap(Epf::getMemberId, epf -> epf));

                for(int rowNum = 2; rowNum <= 180; rowNum++)
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

    @Override
    public String createTextFiles(MultipartFile multipartFile) throws IOException
    {
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);

        List<String> staffList = new ArrayList<>();
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            Sheet sheet = workbook.getSheetAt(0);
            for(int rowNum = 2; rowNum <= 180; rowNum++)
            {
                Row row = sheet.getRow(rowNum);
                Cell staffNoCell = row.getCell(5);

                String staffNo;
                if(staffNoCell.getCellType() == CellType.STRING)
                    staffNo = staffNoCell.getStringCellValue();
                else
                    staffNo = String.valueOf((int) staffNoCell.getNumericCellValue());

                staffList.add(staffNo);
            }
        }

        staffList.forEach(staffNo ->
        {
            List<IEpfView> epfViewList = epfRepository.findByStaffNo(staffNo);
            StringBuilder employeeEpfData = new StringBuilder();
            String memberId = epfViewList.get(0).getMemberId();
            Map<String, Integer> finalDueMap = new HashMap<>();
            epfViewList.forEach(epfView ->
            {
                String monthYear = epfView.getMonthYear().substring(0,2) + "/"
                        + epfView.getMonthYear().substring(3);
                double rate = epfView.getRate();
                Long epsDue = epfView.getEpsDue();
                double finalEpsDue = epsDue + ((rate/12 * epsDue) / 100);
                employeeEpfData.append(epfView.getMemberId())
                        .append(SEPARATOR)
                        .append(monthYear)
                        .append(SEPARATOR);
                if(monthYear.equalsIgnoreCase("11/1995"))
                {
                    Integer finalDue = (int) Math.round(finalEpsDue / 2);
                    finalDueMap.put(epfView.getMonthYear(), finalDue);
                    employeeEpfData
                            .append(Math.round(epfView.getTotalSalary()/2))
                            .append(SEPARATOR)
                            .append(Math.round((float) epfView.getEpsOnTotalWage() / 2))
                            .append(SEPARATOR)
                            .append(Math.round((float) epfView.getAdditionalContribution() / 2))
                            .append(SEPARATOR)
                            .append(Math.round((float) epfView.getEpsDeposited() / 2))
                            .append(SEPARATOR)
                            .append(finalDue);
                }
                else
                {
                    Integer finalDue = (int) Math.round(finalEpsDue);
                    finalDueMap.put(epfView.getMonthYear(), finalDue);
                    employeeEpfData.append((int)epfView.getTotalSalary())
                            .append(SEPARATOR)
                            .append(epfView.getEpsOnTotalWage())
                            .append(SEPARATOR)
                            .append(epfView.getAdditionalContribution())
                            .append(SEPARATOR)
                            .append(epfView.getEpsDeposited())
                            .append(SEPARATOR)
                            .append(Math.round(finalEpsDue));
                }
                employeeEpfData.append(SEPARATOR)
                        .append(rate)
                        .append("\n");
            });
            epfCustomRepository.updateFinalDue(staffNo, finalDueMap);

            removeLastLine(employeeEpfData);

            try (FileOutputStream fos = new FileOutputStream(outputPath + "/Employee/"
                    + memberId + ".txt"))
            {
                byte[] bytes = employeeEpfData.toString().getBytes(StandardCharsets.UTF_8);

                fos.write(bytes);
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        });

        return "Files created successfully.";
    }

    @Override
    public String createTextFilesMulti(MultipartFile multipartFile) throws IOException
    {
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);

        Set<Staff> staffList = new HashSet<>();
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            Sheet sheet = workbook.getSheetAt(0);
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");

            for(int rowNum = 1; rowNum <= 27; rowNum++)
            {
                Row row = sheet.getRow(rowNum);
                Cell staffNoCell = row.getCell(2);
                Cell dateOfLeavingCell = row.getCell(4);
                // Parse to LocalDate
                LocalDate localDate = LocalDate.parse(dateOfLeavingCell.getStringCellValue(), formatter);

                // Convert to Date
                Date dateOfLeaving = Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());

                String staffNo;
                if(staffNoCell.getCellType() == CellType.STRING)
                    staffNo = staffNoCell.getStringCellValue();
                else
                    staffNo = String.valueOf((int) staffNoCell.getNumericCellValue());

                Staff staff = new Staff();
                staff.setStaffNumber(staffNo);
                staff.setDateOfLeaving(dateOfLeaving);
                staffList.add(staff);
            }
        }

        staffList.forEach(staff ->
        {
            List<IEpfView> epfViewList = epfRepository.findByStaffNo(staff.getStaffNumber());
                StringBuilder employeeEpfData = new StringBuilder();
                String memberId = epfViewList.get(0).getMemberId();
                Map<String, Integer> finalDueMap = new HashMap<>();
                epfViewList.forEach(epfView ->
                {
                    String monthYearStr = epfView.getMonthYear();
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/yyyy");

                    // Parse to YearMonth
                    YearMonth yearMonth = YearMonth.parse(monthYearStr, formatter);

                    // Convert to Date (set day to the first of the month)
                    Date date = Date.from(yearMonth.atDay(1).atStartOfDay(ZoneId.systemDefault()).toInstant());

                    log.info(monthYearStr);
                    if(monthYearStr.equals("10/2016"))
                    {
                        log.info("hey");
                    }
                    if(date.before(staff.getDateOfLeaving()))
                    {
                        String monthYear = epfView.getMonthYear().substring(0,2) + "/"
                                + epfView.getMonthYear().substring(3);
                        double rate = epfView.getRate();
                        Long epsDue = epfView.getEpsDue();
                        double finalEpsDue = epsDue + ((rate/12 * epsDue) / 100);
                        employeeEpfData.append(epfView.getMemberId())
                                .append(SEPARATOR)
                                .append(monthYear)
                                .append(SEPARATOR);
                        Integer finalDue = (int) Math.round(finalEpsDue);
                        finalDueMap.put(epfView.getMonthYear(), finalDue);
                        employeeEpfData.append((int)epfView.getTotalSalary())
                                .append(SEPARATOR)
                                .append(epfView.getEpsOnTotalWage())
                                .append(SEPARATOR)
                                .append(epfView.getAdditionalContribution())
                                .append(SEPARATOR)
                                .append(epfView.getEpsDeposited())
                                .append(SEPARATOR)
                                .append(Math.round(finalEpsDue));
                        employeeEpfData.append(SEPARATOR)
                                .append(rate)
                                .append("\n");
                    }

                });
                epfCustomRepository.updateFinalDue(staff.getStaffNumber(), finalDueMap);

                removeLastLine(employeeEpfData);

                try (FileOutputStream fos = new FileOutputStream(outputPath + "/Employee/"
                        + memberId + ".txt"))
                {
                    byte[] bytes = employeeEpfData.toString().getBytes(StandardCharsets.UTF_8);

                    fos.write(bytes);
                }
                catch (IOException e)
                {
                    e.printStackTrace();
                }

        });

        return "Files created successfully.";
    }

    public String getRatePercentage(String monthYear)
    {
        int month = Integer.parseInt(monthYear.substring(0,2));
        int year = Integer.parseInt(monthYear.substring(3));

        if(year>=1995 && year<=2000 || (year == 2001 && month >= 1 && month <= 6))
        {
            return "12";
        }
        else if((year == 2001 && month>=7 && month <=12) || (year==2002 && month==1))
        {
            return "11";
        }
        else if((year>=2002 && year<=2005) || (year==2006 && month==1)
                || (year==2011 && month>=2) || (year==2012 && month<=1))
        {
            return "9.5";
        }
        else if(year >= 2006 && year <= 2010 || (year == 2011 && month==1)
        || (year == 2013 && month>=2) ||  (year == 2014 && month==1))
        {
            return "8.5";
        }
        else if((year == 2012) ||  (year == 2013 && month==1))
        {
            return "8.25";
        }

        else if((year >= 2014 && year <=2015) ||  (year == 2016 && month==1))
        {
            return "8.75";
        }

        else if(year >= 2016)
        {
            return "8.8";
        }
        return "0";
    }

    /*@Override
    public String createExcelFile(MultipartFile multipartFile) throws IOException
    {
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);

        List<String> staffList = new ArrayList<>();
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            Sheet sheet = workbook.getSheetAt(0);
            for(int rowNum = 2; rowNum <= 180; rowNum++)
            {
                Row row = sheet.getRow(rowNum);
                Cell staffNoCell = row.getCell(5);

                String staffNo;
                if(staffNoCell.getCellType() == CellType.STRING)
                    staffNo = staffNoCell.getStringCellValue();
                else
                    staffNo = String.valueOf((int) staffNoCell.getNumericCellValue());

                staffList.add(staffNo);
            }
        }

        FileInputStream fileInputStream1 = new FileInputStream(outputPath + "/" + "Employee Wise.xlsx");
        try(Workbook workbook = WorkbookFactory.create(fileInputStream1))
        {
            Sheet sheet = workbook.getSheetAt(0);
            AtomicInteger rowNum = new AtomicInteger(2);
            CellStyle cellStyle = workbook.createCellStyle();
            createCellStyle(cellStyle, workbook);
            staffList.forEach(staffNo ->
            {
                List<Epf> epfViewList = epfRepository.findAllByStaffNo(staffNo);
                epfViewList.forEach(epfView ->
                {
                    Row row = sheet.createRow(rowNum.get());
                    rowNum.set(rowNum.get()+1);

                    Cell sNoCell = row.createCell(0);
                    sNoCell.setCellValue(epfView.getSNo());
                    sNoCell.setCellStyle(cellStyle);

                    Cell slipNoCell = row.createCell(1);
                    slipNoCell.setCellValue(epfView.getSlipNo());
                    slipNoCell.setCellStyle(cellStyle);

                    Cell memberIdCell = row.createCell(2);
                    memberIdCell.setCellValue(epfView.getMemberId());
                    memberIdCell.setCellStyle(cellStyle);

                    Cell nameCell = row.createCell(3);
                    nameCell.setCellValue(epfView.getName());
                    nameCell.setCellStyle(cellStyle);

                    Cell categoryCell = row.createCell(4);
                    categoryCell.setCellValue(epfView.getCategory());
                    categoryCell.setCellStyle(cellStyle);

                    Cell staffNoCell = row.createCell(5);
                    staffNoCell.setCellValue(epfView.getStaffNo());
                    staffNoCell.setCellStyle(cellStyle);

                    Cell basicCell = row.createCell(6);
                    basicCell.setCellValue(epfView.getBasic());
                    basicCell.setCellStyle(cellStyle);

                    Cell daVdaCell = row.createCell(7);
                    daVdaCell.setCellValue(epfView.getDaVda());
                    daVdaCell.setCellStyle(cellStyle);

                    Cell ppCell = row.createCell(8);
                    ppCell.setCellValue(epfView.getPp());
                    ppCell.setCellStyle(cellStyle);

                    Cell basicArrearCell = row.createCell(9);
                    basicArrearCell.setCellValue(epfView.getBasicArrears());
                    basicArrearCell.setCellStyle(cellStyle);

                    Cell daArrearCell = row.createCell(10);
                    daArrearCell.setCellValue(epfView.getDaArrears());
                    daArrearCell.setCellStyle(cellStyle);

                    Cell ppArrearCell = row.createCell(11);
                    ppArrearCell.setCellValue(epfView.getPpArrears());
                    ppArrearCell.setCellStyle(cellStyle);

                    Cell cpfArrearCell = row.createCell(12);
                    cpfArrearCell.setCellValue(epfView.getCpfArrears());
                    cpfArrearCell.setCellStyle(cellStyle);

                    Cell epsArrearCell = row.createCell(13);
                    epsArrearCell.setCellValue(epfView.getEpsArrears());
                    epsArrearCell.setCellStyle(cellStyle);

                    Cell totalCell = row.createCell(14);
                    totalCell.setCellValue(epfView.getTotalSalary());
                    totalCell.setCellStyle(cellStyle);

                    Cell epsDepositedCell = row.createCell(15);
                    epsDepositedCell.setCellValue(epfView.getEpsDeposited());
                    epsDepositedCell.setCellStyle(cellStyle);

                    Cell epsOnTotalWageCell = row.createCell(16);
                    epsOnTotalWageCell.setCellValue(epfView.getEpsOnTotalWage());
                    epsOnTotalWageCell.setCellStyle(cellStyle);

                    Cell additionalContributionCell = row.createCell(17);
                    additionalContributionCell.setCellValue(epfView.getAdditionalContribution());
                    additionalContributionCell.setCellStyle(cellStyle);

                    Cell epsDueCell = row.createCell(18);
                    epsDueCell.setCellValue(epfView.getEpsDue());
                    epsDueCell.setCellStyle(cellStyle);

                    Cell monthYearCell = row.createCell(19);
                    monthYearCell.setCellValue(epfView.getMonthYear());
                    monthYearCell.setCellStyle(cellStyle);
                });
            });

            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + "Employee Wise1.xlsx");
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return "File created successfully.";
    }*/

    @Override
    public String createExcelFile(MultipartFile multipartFile) throws IOException
    {
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            CellStyle cellStyle = workbook.createCellStyle();
            createCellStyle(cellStyle, workbook);
            CellStyle headerCellStyle = workbook.createCellStyle();
            createHeaderCellStyle(headerCellStyle, workbook);
            for(int i=0;i<=254;i++)
            {
                Sheet sheet = workbook.getSheetAt(i);
                Row header = sheet.getRow(1);
                Cell rateHeader = header.createCell(18);
                rateHeader.setCellStyle(headerCellStyle);
                rateHeader.setCellValue("Rate (%)");

                Cell finalDueHeader = header.createCell(19);
                finalDueHeader.setCellValue("Final Due");
                finalDueHeader.setCellStyle(headerCellStyle);
                for(int rowNum = 2; rowNum <= 180; rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    Cell memberIdCell = row.getCell(2);
                    String memberId = memberIdCell.getStringCellValue();

                    List<Epf> epfList = epfRepository
                            .findByMemberIdAndMonthYear(List.of(memberId),
                                    sheet.getSheetName());

                    Epf epf = epfList.get(0);
                    Cell rateCell = row.createCell(18);
                    rateCell.setCellStyle(headerCellStyle);
                    rateCell.setCellValue(epf.getRate());
                    Cell finalDueCell = row.createCell(19);
                    finalDueCell.setCellStyle(headerCellStyle);
                    finalDueCell.setCellValue(epf.getFinalDue());
                }
            }

            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + "Final Due.xlsx");
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return "File created successfully.";
    }

    private void removeLastLine(StringBuilder stringBuilder) {
        int lastNewLineIndex = stringBuilder.lastIndexOf("\n");

        if (lastNewLineIndex >= 0) {
            stringBuilder.delete(lastNewLineIndex, stringBuilder.length());
        }
    }


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
                for(int rowNum = 2; rowNum <= 180; rowNum++)
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

                for(int rowNum = 2; rowNum <= 180; rowNum++)
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

    private Epf createEmployee(Row row, String month, FormulaEvaluator evaluator)
    {
        CellDto cellDto = getCellDto(row, month, evaluator);

        log.info("Row: " + cellDto);

        return Epf.builder()
                .sNo(cellDto.getSNo())
                .slipNo(cellDto.getSlipNo())
                .memberId(cellDto.getMemberId())
                .name(cellDto.getName())
                .category(cellDto.getCategory())
                .staffNo(cellDto.getStaffNo())
                .basic(cellDto.getBasic())
                .pp(cellDto.getPp())
                .daVda(cellDto.getDaVda())
                .basicArrears(cellDto.getBasicArrears())
                .ppArrears(cellDto.getPpArrears())
                .daArrears(cellDto.getDaArrears())
                .cpfArrears(cellDto.getCpfArrears())
                .epsArrears(cellDto.getEpsArrears())
                .totalSalary(cellDto.getTotalSalary())
                .monthYear(month)
                .epsDeposited(cellDto.getEpsDeposited())
                .epsOnTotalWage(cellDto.getEpsOnTotalWage())
                .additionalContribution(cellDto.getAdditionalContribution())
                .epsDue(cellDto.getEpsDue())
                .rate(cellDto.getRate())
                .build();
    }

    private Epf createEmployeeMulti(Row row, String memberId, String staffNumber)
    {
        Date date = row.getCell(1).getDateCellValue();
        LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();

        Month monthObj = localDate.getMonth();
        int monthValue = monthObj.getValue();
        String monthStr = String.valueOf(monthValue);
        if(monthValue<10)
            monthStr = "0".concat(monthStr);
        String year = String.valueOf(localDate.getYear());

        String month = monthStr.concat("/").concat(year);

        if(month.equals("11/2011"))
        {
            log.info("hey");
        }
        String rate = getRatePercentage(month);

        CellDto cellDto = getCellDtoMulti(row, month);

        log.info("Row: " + cellDto);

        return Epf.builder()
                .sNo(cellDto.getSNo())
                .slipNo(0)
                .staffNo(staffNumber)
                .basic(0.0)
                .pp(0.0)
                .daVda(0.0)
                .basicArrears(0.0)
                .ppArrears(0.0)
                .category("")
                .daArrears(0.0)
                .cpfArrears(0.0)
                .epsArrears(0.0)
                .finalDue(0L)
                .rate(Double.parseDouble(rate))
                .name("")
                .memberId(memberId)
                .monthYear(month)
                .totalSalary(cellDto.getTotalSalary())
                .epsOnTotalWage(cellDto.getEpsOnTotalWage())
                .additionalContribution(cellDto.getAdditionalContribution())
                .epsDeposited(cellDto.getEpsDeposited())
                .epsDue(cellDto.getEpsDue())
                .build();
    }

    private void verifyEmployee(Row row, Epf epf, String month, FormulaEvaluator evaluator,
                                List<Epf> epfMismatchList)
    {
        CellDto cellDto = getCellDto(row, month, evaluator);
        Epf epfMismatch = new Epf();
        epfMismatch.setStaffNo(epf.getStaffNo());
        epfMismatch.setName(epf.getName());

        //For any data mismatch that happens
        AtomicBoolean mismatch = new AtomicBoolean(false);

        //Perform verification for all cell data in a row
        dataValidator.verifySnoSlipMemberNameCatStaff(epf, epfMismatch, cellDto, mismatch);
        dataValidator.verifyBasicPpDaBasicArrPpArrDaArr(epf, epfMismatch, cellDto, mismatch);
        dataValidator.verifyCpfArrEpsArrTotalSalMonth(epf, epfMismatch, cellDto, mismatch, month);
        dataValidator.verifyEpsDepEpsTotWageAddContEpsDue(epf, epfMismatch, cellDto, mismatch);

        if(mismatch.get())
            epfMismatchList.add(epfMismatch);
    }

    private CellDto getCellDto(Row row, String month, FormulaEvaluator evaluator)
    {
        String[] monthArr = month.split("-");
        int mon = Integer.parseInt(monthArr[0]);
        int year = Integer.parseInt(monthArr[1]);

        //Read all 18 cells in a row
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
        Cell totalSalaryCell = row.getCell(14);
        double totalSalary = 0;
        if(totalSalaryCell!=null) {
            if (evaluator.evaluate(row.getCell(14)) != null) {
                totalSalary = Math.round(evaluator.evaluate(row.getCell(14)).getNumberValue());
            }
        }
        Cell epsDeposited = row.getCell(15);
        Cell epsOnTotalWage = row.getCell(16);
        Cell epsDue;
        Cell rate;
        long additionalContribution = 0L;
        if(year >= 2015 || (year == 2014 && mon>=9))
        {
            Cell additionalContributionCell = row.getCell(17);
            if(additionalContributionCell != null)
            {
                additionalContribution = (long) additionalContributionCell.getNumericCellValue();
            }
            epsDue = row.getCell(18);
            rate = row.getCell(20);
        }
        else
        {
            epsDue = row.getCell(17);
            rate = row.getCell(18);
        }

        String staffCellNoStr = "";
        if(staffNo != null) {
            if (staffNo.getCellType() == CellType.STRING)
                staffCellNoStr = staffNo.getStringCellValue();
            else
                staffCellNoStr = String.valueOf((int) staffNo.getNumericCellValue());
        }

        return CellDto.builder()
                .sNo(sNo != null ? (int)sNo.getNumericCellValue():0)
                .slipNo(slipNo!=null?(int)slipNo.getNumericCellValue():0)
                .memberId(memberId != null ? memberId.getStringCellValue():"")
                .name(name != null ? name.getStringCellValue():"")
                .category(category != null ? category.getStringCellValue():"")
                .staffNo(staffCellNoStr)
                .basic(basic != null ? basic.getNumericCellValue():0)
                .pp(pp != null ? pp.getNumericCellValue():0)
                .daVda(daVda != null ? daVda.getNumericCellValue():0)
                .basicArrears(basicArrears != null ? basicArrears.getNumericCellValue():0)
                .ppArrears(ppArrears != null ? ppArrears.getNumericCellValue():0)
                .daArrears(daArrears != null ? daArrears.getNumericCellValue():0)
                .cpfArrears(cpfArrears != null ? cpfArrears.getNumericCellValue():0)
                .epsArrears(epsArrears != null ? epsArrears.getNumericCellValue():0)
                .totalSalary(totalSalary)
                .epsDeposited(epsDeposited != null ? (long) epsDeposited.getNumericCellValue():0L)
                .epsOnTotalWage(epsOnTotalWage != null ? (long) epsOnTotalWage.getNumericCellValue():0L)
                .additionalContribution(additionalContribution)
                .epsDue(epsDue != null ? (long) epsDue.getNumericCellValue():0L)
                .rate(rate != null ? rate.getNumericCellValue() :0)
                .build();
    }

    private CellDto getCellDtoMulti(Row row, String month)
    {
        log.info("For month:{}", month);

        String[] monthArr = month.split("/");
        int monthVal = Integer.parseInt(monthArr[0]);
        int year = Integer.parseInt(monthArr[1]);
        //Read all 18 cells in a row
        Cell sNo = row.getCell(0);
        Cell totalSalaryCell = row.getCell(2);
        Cell epsDepositedCell = row.getCell(5);
        if(year==2015 && monthVal==1)
        {
            log.info("eps 8.33: {}");
        }
        long totalSalary = totalSalaryCell != null ?(long) totalSalaryCell.getNumericCellValue() : 0;
        long additionalContribution = 0L;
        if(totalSalary>=15000 && (year==2014 && monthVal>=9 || year>2014))
        {
            additionalContribution = Math.round((totalSalary - 15000) * 0.0116);
        }

        long epsDeposited = epsDepositedCell != null ? (long) epsDepositedCell.getNumericCellValue() : 0L;
        long epsOnTotalWage = Math.round((8.33 * totalSalary)/100.0);
        long epsDue = 0L;
        if(epsDeposited < epsOnTotalWage) {
            epsDue = epsOnTotalWage + additionalContribution - epsDeposited;
        }

        return CellDto.builder()
                .sNo(sNo != null ? (int)sNo.getNumericCellValue():0)
                .totalSalary((double)totalSalary)
                .epsDeposited(epsDeposited)
                .epsOnTotalWage(epsOnTotalWage)
                .additionalContribution(additionalContribution)
                .epsDue(epsDue)
                .build();
    }
}
