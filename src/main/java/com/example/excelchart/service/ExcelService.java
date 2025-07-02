package com.example.excelchart.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Service
public class ExcelService {

    public List<CampaignData> loadCampaignData(File excelFile) throws Exception {
        List<CampaignData> campaigns = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile))) {
            Sheet sheet = workbook.getSheet("1. campaign");
            if (sheet == null) return campaigns;
            int headerRowIndex = 3;
            for (int i = headerRowIndex + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String campaignId = getStringCellValue(row.getCell(1));
                if (campaignId == null || campaignId.trim().isEmpty()) continue;
                CampaignData campaign = new CampaignData(
                    campaignId,
                    getStringCellValue(row.getCell(2)),
                    getDoubleCellValue(row.getCell(3)),
                    getLongCellValue(row.getCell(4)),
                    getDoubleCellValue(row.getCell(5)),
                    getBigDecimalCellValue(row.getCell(6)),
                    getBigDecimalCellValue(row.getCell(7)),
                    getBigDecimalCellValue(row.getCell(8)),
                    getBigDecimalCellValue(row.getCell(9)),
                    getBigDecimalCellValue(row.getCell(10)),
                    getBigDecimalCellValue(row.getCell(11)),
                    getBigDecimalCellValue(row.getCell(12)),
                    getBigDecimalCellValue(row.getCell(13))
                );
                campaigns.add(campaign);
            }
        }
        return campaigns;
    }

    public List<DateData> loadDateData(File excelFile) throws Exception {
        List<DateData> dates = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile))) {
            Sheet sheet = workbook.getSheet("3. date");
            if (sheet == null) return dates;
            int headerRowIndex = 3;
            DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("M/d/yyyy");
            for (int i = headerRowIndex + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String campaignId = getStringCellValue(row.getCell(1));
                if (campaignId == null || campaignId.trim().isEmpty()) continue;
                String dateString = getStringCellValue(row.getCell(2));
                LocalDate date = (dateString != null && !dateString.isEmpty())
                        ? LocalDate.parse(dateString, dateFormatter) : null;
                // Add other fields as needed
                DateData dateData = new DateData(
                    campaignId,
                    date
                    // ... other fields
                );
                dates.add(dateData);
            }
        }
        return dates;
    }

    // --- Helper methods ---
    private static String getStringCellValue(Cell cell) {
        if (cell == null) return null;
        cell.setCellType(CellType.STRING);
        String value = cell.getStringCellValue();
        return value.trim().isEmpty() ? null : value.trim();
    }

    private static Double getDoubleCellValue(Cell cell) {
        if (cell == null) return null;
        try {
            return cell.getNumericCellValue();
        } catch (Exception e) {
            if (cell.getCellType() == CellType.STRING) {
                try {
                    return Double.parseDouble(cell.getStringCellValue());
                } catch (NumberFormatException ex) {
                    return null;
                }
            }
            return null;
        }
    }

    private static Long getLongCellValue(Cell cell) {
        if (cell == null) return null;
        try {
            return (long) cell.getNumericCellValue();
        } catch (Exception e) {
            if (cell.getCellType() == CellType.STRING) {
                try {
                    return Long.parseLong(cell.getStringCellValue().split("\\.")[0]);
                } catch (NumberFormatException ex) {
                    return null;
                }
            }
            return null;
        }
    }

    private static BigDecimal getBigDecimalCellValue(Cell cell) {
        if (cell == null) return null;
        try {
            return BigDecimal.valueOf(cell.getNumericCellValue());
        } catch (Exception e) {
            if (cell.getCellType() == CellType.STRING) {
                try {
                    return new BigDecimal(cell.getStringCellValue());
                } catch (NumberFormatException ex) {
                    return null;
                }
            }
            return null;
        }
    }
}