package com.hcl.ems.services.impl;

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
import java.util.List;

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
}
