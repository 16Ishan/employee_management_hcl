package com.hcl.ems.services.impl;

import com.hcl.ems.dto.BasicDto;
import com.hcl.ems.services.SheetService;
import com.hcl.ems.utils.CommonUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

@Service
@Slf4j
public class SheetServiceImpl implements SheetService
{
    @Value("${spring.config.location}")
    private String outputPath;
    private static final String OUTPUT_FILE = "Salary Records.xlsx";

    @Override
    public String deleteColumnsFromMultipleSheets(String startDate, String endDate,
                                                  MultipartFile multipartFile) throws IOException
    {
        List<String> sheetList = CommonUtil.getSheetList(startDate, endDate);
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
    public String missingTotalFormula(String startDate, String endDate,
                                      MultipartFile multipartFile, String action) throws IOException
    {
        List<String> sheetList = CommonUtil.getSheetList(startDate, endDate);
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            for(String month:sheetList)
            {
                if(action.equals("check"))
                    checkMissingFormula(month, workbook);
                else
                    addMissingFormula(month, workbook);
            }
            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + OUTPUT_FILE);
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return "Logged missing formulas successfully";
    }

    public void checkMissingFormula(String month, Workbook workbook)
    {
        List<String> missingMemberFormula = new ArrayList<>();
        Sheet sheet = workbook.getSheet(month);

        log.info("***********"+month+"************");
        for(int rowNum = 2; rowNum <= 180; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            Cell totalCell = row.getCell(14);
            if(!(totalCell != null && totalCell.getCellType() == CellType.FORMULA))
            {
                Cell staffNoCell = row.getCell(5);
                if(staffNoCell.getCellType() == CellType.STRING)
                    missingMemberFormula.add(staffNoCell.getStringCellValue());
                else
                    missingMemberFormula.add(String.valueOf(
                            (int)staffNoCell.getNumericCellValue()));
            }
        }
        log.info("Missing formula for "+missingMemberFormula+"\n");
    }

    public void addMissingFormula(String month, Workbook workbook)
    {
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Sheet sheet = workbook.getSheet(month);
        String formulaPattern = "SUM(G%s:N%s)";

        CellStyle cellStyle = workbook.createCellStyle();
        createCellStyle(cellStyle, workbook);

        for(int rowNum = 2; rowNum <= 180; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            Cell totalCell = row.getCell(14);
            if(!(totalCell != null && totalCell.getCellType() == CellType.FORMULA))
            {
                Cell cell = row.createCell(14);
                String formula = String.format(formulaPattern, rowNum + 1, rowNum + 1);
                cell.setCellFormula(formula);
                cell.setCellStyle(cellStyle);
                formulaEvaluator.evaluateFormulaCell(cell);
            }
        }
    }

    @Override
    public String removeMergedCells(String startDate, String endDate,
                                    MultipartFile multipartFile) throws IOException
    {
        List<String> sheetList = CommonUtil.getSheetList(startDate, endDate);
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            for(String month:sheetList)
            {
                removeMerge(month, workbook, 15, 17, 1);
                removeMerge(month, workbook, 18, 20, 1);
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

    @Override
    public String addEps(String startDate, String endDate, MultipartFile multipartFile) throws IOException
    {
        List<String> sheetList = CommonUtil.getSheetList(startDate, endDate);
        File file = CommonUtil.convertMultiPartToFile(multipartFile, outputPath);

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            for(String month:sheetList)
            {
                addEpsColumns(month, workbook);
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

    @Override
    public Map<String, List<String>> addMissingSalary(String startDate, String endDate, List<BasicDto> basicDtoList) throws IOException
    {
        List<String> sheetList = CommonUtil.getSheetList(startDate, endDate);
        File file = new File("D:\\Downloads\\Salary Records1.xlsx");

        Map<String, List<String>> successMonthMap = new HashMap<>();

        FileInputStream fileInputStream = new FileInputStream(file);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream))
        {
            basicDtoList.forEach(basicDto -> {
                List<Double> basicList = basicDto.getBasicList();
                String memberId = "UPALD00187490000000" + basicDto.getMemberId();
                List<String> successMonths = new ArrayList<>();
                for(int i = 0; i<sheetList.size(); i++)
                {
                    boolean result = addBasic(sheetList.get(i), basicList.get(i),
                            memberId, workbook);
                    if(result)
                        successMonths.add(sheetList.get(i));
                    else
                        log.error(memberId);
                }

                successMonthMap.put(memberId, successMonths);
            });

            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + OUTPUT_FILE);
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return successMonthMap;
    }

    @Override
    public String addMissingMonths(String startDate, String endDate,
                                                      MultipartFile multipartFile1,
                                                      MultipartFile multipartFile2,
                                                      MultipartFile multipartFile3) throws IOException
    {
        List<String> sheetList = CommonUtil.getSheetList(startDate, endDate);
        Map<Integer, Map<String, Double>> employeeSalaryMap = new HashMap<>();

        File file1 = CommonUtil.convertMultiPartToFile(multipartFile1, outputPath);
        File file2 = CommonUtil.convertMultiPartToFile(multipartFile2, outputPath);
        File file3 = CommonUtil.convertMultiPartToFile(multipartFile3, outputPath);

        FileInputStream fileInputStream1 = new FileInputStream(file1);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream1))
        {
            extractAndBuildSalaryMap(employeeSalaryMap, workbook.getSheetAt(0));
        }

        FileInputStream fileInputStream2 = new FileInputStream(file2);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream2))
        {
            extractAndBuildSalaryMap(employeeSalaryMap, workbook.getSheetAt(0));
        }

        FileInputStream fileInputStream3 = new FileInputStream(file3);
        try(Workbook workbook = WorkbookFactory.create(fileInputStream3))
        {
            for(String month:sheetList)
            {
                addMissingData(employeeSalaryMap, workbook, month);
            }
            try {
                FileOutputStream fos = new FileOutputStream(outputPath + "/" + OUTPUT_FILE);
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return "Data added";
    }

    private void addMissingData(Map<Integer, Map<String, Double>> employeeSalaryMap, Workbook workbook, String month)
    {
        List<Integer> missingEmployee = new ArrayList<>();
        Sheet sheet = workbook.getSheet(month);

        CellStyle cellStyle = workbook.createCellStyle();
        createCellStyle(cellStyle, workbook);

        for(int rowNum = 2; rowNum <= 180; rowNum++)
        {
            Row row = sheet.getRow(rowNum);

            Cell cell1 = row.createCell(6);
            cell1.setCellStyle(cellStyle);
            cell1.setCellValue(0);

            Cell cell2 = row.createCell(7);
            cell2.setCellStyle(cellStyle);
            cell2.setCellValue(0);

            Cell cell3 = row.createCell(8);
            cell3.setCellStyle(cellStyle);
            cell3.setCellValue(0);

            Cell cell4 = row.createCell(9);
            cell4.setCellStyle(cellStyle);
            cell4.setCellValue(0);

            Cell cell5 = row.createCell(10);
            cell5.setCellStyle(cellStyle);
            cell5.setCellValue(0);

            Cell cell6 = row.createCell(11);
            cell6.setCellStyle(cellStyle);
            cell6.setCellValue(0);

            Cell cell7 = row.createCell(12);
            cell7.setCellStyle(cellStyle);
            cell7.setCellValue(0);

            Cell staffNoCell = row.getCell(5);
            int staffNo = -1;
            if(staffNoCell != null)
            {
                if(staffNoCell.getCellType().equals(CellType.STRING))
                    staffNo = Integer.parseInt(staffNoCell.getStringCellValue()
                        .substring(3).concat("11111"));
                else
                    staffNo = (int)staffNoCell.getNumericCellValue();
            }

            Map<String, Double> salaryMap = employeeSalaryMap.get(staffNo);
            if(salaryMap == null)
            {
                missingEmployee.add(staffNo);
                continue;
            }
            Double totalSalary = salaryMap.get(month);

            cell1.setCellValue(totalSalary);
        }

        log.info(missingEmployee.toString());
    }

    public static void createCellStyle(CellStyle cellStyle, Workbook workbook)
    {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        Font font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    public static void createHeaderCellStyle(CellStyle cellStyle, Workbook workbook)
    {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        Font font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    private void extractAndBuildSalaryMap(Map<Integer, Map<String, Double>> employeeSalaryMap, Sheet sheet)
    {
        try {
            for(int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++)
            {
                Row row = sheet.getRow(rowNum);
                Cell monthYearCell = row.getCell(0);
                if(monthYearCell == null)
                {
                    continue;
                }
                Cell staffNoCell = row.getCell(1);
                Cell totalSalaryCell = row.getCell(10);

                String monthYear = String.valueOf((int)monthYearCell.getNumericCellValue());
                monthYear =  monthYear.substring(4) + "-" + monthYear.substring(0,4);
                Integer staffNo = (int) staffNoCell.getNumericCellValue();
                Double totalSalary = totalSalaryCell.getNumericCellValue();

                Map<String, Double> salaryMap = employeeSalaryMap.get(staffNo);
                if(salaryMap != null && !salaryMap.isEmpty())
                {
                    salaryMap.put(monthYear, totalSalary);
                }
                else
                {
                    salaryMap = new HashMap<>();
                    salaryMap.put(monthYear, totalSalary);
                }

                employeeSalaryMap.put(staffNo, salaryMap);
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

    }

    private boolean addBasic(String month, Double basic, String memberId, Workbook workbook)
    {
        Sheet sheet = workbook.getSheet(month);

        for(int rowNum = 2; rowNum <= 180; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            Cell memberIdCell = row.getCell(2);
            String memId = memberIdCell.getStringCellValue();

            if(memberId.equalsIgnoreCase(memId))
            {
                Cell basicCell = row.getCell(6);
                basicCell.setCellValue(basic);
                return true;
            }
        }
        return false;
    }

    public void deleteColumn(String month, Workbook workbook)
    {
        Sheet sheet = workbook.getSheet(month);
        Row row1 = sheet.getRow(0);
        Cell cell1= row1.getCell(1);
        if (cell1 != null) {
            row1.removeCell(cell1);
        }
        for(int rowNum = 0; rowNum <= 180; rowNum++)
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

    public void removeMerge(String month, Workbook workbook, int start, int end, int region)
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
    public void addEpsColumns(String month, Workbook workbook)
    {
        Sheet sheet = workbook.getSheet(month);
        sheet.setColumnWidth(15, 7268);

        CellStyle cell1Style = workbook.createCellStyle();
        CellStyle cell2Style = workbook.createCellStyle();
        CellStyle cell3Style = workbook.createCellStyle();
        CellStyle cell4Style = workbook.createCellStyle();

        createCellStyles(cell1Style, cell2Style, cell3Style, cell4Style, workbook);

        Row row1 = sheet.getRow(0);
        Row row2 = sheet.getRow(1);
        formatRows(row1, row2, cell2Style);

        String[] monthArr = month.split("-");
        int mon = Integer.parseInt(monthArr[0]);
        int year = Integer.parseInt(monthArr[1]);

        Cell cell1 = row1.createCell(15);
        Cell cell2 = row2.createCell(15);
        Cell cell3 = row2.createCell(16);
        Cell cell4;
        Cell cell9;
        if(year >= 2015 || (year == 2014 && mon>=9))
        {
            cell4 = row2.createCell(17);
            cell9 = row2.createCell(18);
            sheet.setColumnWidth(17, 6716);
        }
        else
        {
            cell9 = row2.createCell(17);
            cell4 = null;
        }

        formatCells(cell1, cell2, cell3, cell4, cell9, cell1Style, cell2Style, List.of(mon,year));

        for(int rowNum = 2; rowNum <= 180; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            Cell rateCell = row.getCell(18);
            double rate = rateCell.getNumericCellValue();
            Cell totalSalaryCell = row.getCell(14);
            double totalSalary = Math.round(totalSalaryCell!=null?totalSalaryCell.getNumericCellValue():0);

            double eps;
            double additionalContribution = 0;
            if(year<=2000 || (year == 2001 && mon <= 5))
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
            else if(year>=2002 && year<=2013 || (year == 2001) || (year == 2014 && mon<=8))
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
            else
            {
                if(totalSalary<15000)
                {
                    eps = Math.round(totalSalary*0.0833);
                }
                else
                {
                    eps = Math.round(15000*0.0833);
                    additionalContribution = Math.round(0.0116*(totalSalary - 15000));
                }
            }
            Cell cell5 = row.createCell(15);
            cell5.setCellValue(eps);
            cell5.setCellStyle(cell3Style);

            Cell cell6 = row.createCell(16);
            long epsOnTotalWage = Math.round(totalSalary * 0.0833);
            cell6.setCellValue(epsOnTotalWage);
            cell6.setCellStyle(cell3Style);

            if(year >= 2015 || (year == 2014 && mon>=9))
            {
                Cell cell7 = row.createCell(17);
                cell7.setCellValue(additionalContribution);
                cell7.setCellStyle(cell3Style);

                Cell cell8 = row.createCell(18);
                cell8.setCellValue(epsOnTotalWage + additionalContribution - eps);
                cell8.setCellStyle(cell3Style);

                Cell cell10 = row.createCell(20);
                cell10.setCellValue(rate);
                cell8.setCellStyle(cell3Style);
            }
            else
            {
                Cell cell8 = row.createCell(17);
                cell8.setCellValue(epsOnTotalWage + additionalContribution - eps);
                cell8.setCellStyle(cell3Style);
            }
        }
    }

    private void formatCells(Cell cell1, Cell cell2, Cell cell3, Cell cell4,
                             Cell cell9, CellStyle cell1Style, CellStyle cell2Style, List<Integer> dateList)
    {
        StringBuilder cell1Value = new StringBuilder("Statutory Ceiling Rs.");
        int mon = dateList.get(0);
        int year = dateList.get(1);

        if(year<=2000 || (year == 2001 && mon <= 5))
        {
            cell1Value.append("5000");
        }
        else if(year>=2002 && year<=2013 || (year == 2001) || (year == 2014 && mon<=8))
        {
            cell1Value.append("6500");
        }
        else
        {
            cell1Value.append("15000");
        }

        cell1.setCellStyle(cell1Style);

        cell2.setCellValue("EPS Deposited");
        cell2.setCellStyle(cell2Style);

        cell3.setCellValue("EPS on Total Wage");
        cell3.setCellStyle(cell2Style);

        if(year >= 2015 || (year == 2014 && mon>=9))
        {
            cell4.setCellValue("Additional Contribution 1.16% of (EPF Wage - 15000)");
            cell4.setCellStyle(cell2Style);
        }

        cell9.setCellValue("EPS Due");
        cell9.setCellStyle(cell2Style);

        cell1.setCellValue(cell1Value.toString());
    }

    private void formatRows(Row row1, Row row2, CellStyle cell2Style)
    {
        short height = 1024;
        row1.setHeight(height);
        row2.setHeight(height);

        for(int i=0;i<=14;i++)
        {
            Cell cell = row2.getCell(i);
            cell2Style.setWrapText(true);
            cell.setCellStyle(cell2Style);
        }
    }

    private void createCellStyles(CellStyle cell1Style, CellStyle cell2Style,
                                  CellStyle cell3Style, CellStyle cell4Style, Workbook workbook)
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

        cell3Style.setBorderTop(BorderStyle.THIN);
        cell3Style.setBorderRight(BorderStyle.THIN);
        cell3Style.setBorderBottom(BorderStyle.THIN);
        cell3Style.setBorderLeft(BorderStyle.THIN);
        Font font3 = workbook.createFont();
        font3.setFontName(fontStyle);
        font3.setFontHeightInPoints((short) 11);
        cell3Style.setFont(font3);
        cell3Style.setAlignment(HorizontalAlignment.CENTER);

        cell4Style.setBorderTop(BorderStyle.THIN);
        cell4Style.setBorderRight(BorderStyle.THIN);
        cell4Style.setBorderBottom(BorderStyle.THIN);
        cell4Style.setBorderLeft(BorderStyle.THIN);
        Font font4 = workbook.createFont();
        font4.setFontName(fontStyle);
        font4.setBold(true);
        font4.setFontHeightInPoints((short) 12);
        cell4Style.setFont(font4);
        cell4Style.setAlignment(HorizontalAlignment.CENTER);
        cell4Style.setVerticalAlignment(VerticalAlignment.TOP);
    }
}
