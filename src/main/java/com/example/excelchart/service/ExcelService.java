package com.example.excelchart.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.excelchart.domain.CampaignData;
import com.example.excelchart.domain.DateData;
import com.example.excelchart.domain.PublisherData;
import com.example.excelchart.domain.CategoryData;
import com.example.excelchart.domain.CityData;
import com.example.excelchart.domain.ScreenData;
import com.example.excelchart.domain.LineItemData;
import com.example.excelchart.domain.CreativeFileData;
import com.example.excelchart.domain.DmaData;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;

@Service
public class ExcelService {

    @Autowired
    private CampaignOverviewPPTService pptService;

    private static final Pattern SHEET_NAME_PATTERN = Pattern.compile("^(\\d+)\\.\\s*(.+)$");
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("M/d/yyyy");
    
    // Separate collections for each data type
    private List<CampaignData> campaignDataList = new ArrayList<>();
    private List<DateData> dateDataList = new ArrayList<>();
    private List<PublisherData> publisherDataList = new ArrayList<>();
    private List<CategoryData> categoryDataList = new ArrayList<>();
    private List<CityData> cityDataList = new ArrayList<>();
    private List<ScreenData> screenDataList = new ArrayList<>();
    private List<LineItemData> lineItemDataList = new ArrayList<>();
    private List<CreativeFileData> creativeFileDataList = new ArrayList<>();
    private List<DmaData> dmaDataList = new ArrayList<>();

    public void readAndGenerateCharts(File excelFile) throws Exception {
        
        // Clear previous data
        clearAllData();
        
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile))) {
            int sheetCount = workbook.getNumberOfSheets();
            
            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                
                processSheet(sheet, sheetName);
            }
        }
        
        // Generate charts with all the collected data
        generateCharts();
    }

    private void processSheet(Sheet sheet, String sheetName) throws Exception {
        // Extract the base name without the number prefix
        String baseName = extractBaseName(sheetName);
        
        // Find the header row (usually row 3 or 4)
        int headerRowIndex = findHeaderRow(sheet);
        if (headerRowIndex == -1) {
            return;
        }
        
        // Get column mappings based on sheet type
        Map<String, Integer> columnMap = getColumnMappings(sheet, headerRowIndex, baseName);
        if (columnMap.isEmpty()) {
            return;
        }
        
        // Process data rows and store in appropriate typed list
        processAndStoreData(sheet, headerRowIndex, columnMap, baseName);
        
    }

    private String extractBaseName(String sheetName) {
        var matcher = SHEET_NAME_PATTERN.matcher(sheetName);
        if (matcher.matches()) {
            return matcher.group(2).trim().toLowerCase();
        }
        return sheetName.toLowerCase();
    }

    private int findHeaderRow(Sheet sheet) {
        // Look for header row (usually row 3 or 4)
        for (int rowIndex = 2; rowIndex <= 5; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                // Check if this row has meaningful headers
                int nonEmptyCells = 0;
                for (int cellIndex = 0; cellIndex < 10; cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    if (cell != null) {
                        String cellValue = getStringCellValue(cell);
                        if (cellValue != null && !cellValue.trim().isEmpty()) {
                            nonEmptyCells++;
                        }
                    }
                }
                if (nonEmptyCells >= 3) {
                    return rowIndex;
                }
            }
        }
        return -1;
    }

    private Map<String, Integer> getColumnMappings(Sheet sheet, int headerRowIndex, String baseName) {
        Map<String, Integer> columnMap = new HashMap<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        
        if (headerRow == null) return columnMap;
        
        // Define expected columns for each sheet type
        Map<String, List<String>> expectedColumns = getExpectedColumns(baseName);
        
        
        for (int cellIndex = 0; cellIndex < headerRow.getLastCellNum(); cellIndex++) {
            Cell cell = headerRow.getCell(cellIndex);
            if (cell != null) {
                String headerValue = getStringCellValue(cell).toLowerCase().trim();
                
                // Check if this header matches any expected column
                for (String expectedColumn : expectedColumns.getOrDefault(baseName, new ArrayList<>())) {
                    if (headerValue.contains(expectedColumn.toLowerCase()) || 
                        expectedColumn.toLowerCase().equals(headerValue)) {
                        columnMap.put(expectedColumn, cellIndex);
                        break;
                    }
                }
            }
        }
        
        return columnMap;
    }

    private Map<String, List<String>> getExpectedColumns(String baseName) {
        Map<String, List<String>> expectedColumns = new HashMap<>();
        
        switch (baseName) {
            case "campaign":
                expectedColumns.put("campaign", Arrays.asList(
                    "Campaign ID", "Campaign", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "date":
                expectedColumns.put("date", Arrays.asList(
                    "Campaign ID", "Date", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "publisher":
                expectedColumns.put("publisher", Arrays.asList(
                    "Campaign ID", "Publisher ID", "Publisher", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "category":
                expectedColumns.put("category", Arrays.asList(
                    "Campaign ID", "Category group", "Category name", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "city":
                expectedColumns.put("city", Arrays.asList(
                    "Campaign ID", "City", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "screen":
                expectedColumns.put("screen", Arrays.asList(
                    "Campaign ID", "Screen ID", "Screen", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", "Latitude", "Longitude",
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "line item":
                expectedColumns.put("line item", Arrays.asList(
                      "Line item", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "creative":
                expectedColumns.put("creative file", Arrays.asList(
                    "Campaign ID", "Creative File ID", "Creative", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
            case "dma":
                expectedColumns.put("dma", Arrays.asList(
                    "Campaign ID", "DMA", "Impressions", "Playouts", 
                    "Imp. / Playout", "Media CPM", "Total CPM", "Media Costs", 
                    "Data Costs", "Platform Costs", "Invoice Amount", "Client Margin", "Total Spend"
                ));
                break;
        }
        
        return expectedColumns;
    }



    private boolean isEmptyRow(Row row) {
        for (int cellIndex = 0; cellIndex < 5; cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            if (cell != null && !getStringCellValue(cell).trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    private Object createDataObject(Row row, Map<String, Integer> columnMap, String baseName) {
        // Get campaign ID first (common field)
        String campaignId = getCellValueAsString(row, columnMap.get("Campaign ID"));
    
        
        
        // Create object based on sheet type
        switch (baseName) {
            case "campaign":
                return createCampaignData(row, columnMap);
            case "date":
                return createDateData(row, columnMap);
            case "publisher":
                return createPublisherData(row, columnMap);
            case "category":
                return createCategoryData(row, columnMap);
            case "city":
                return createCityData(row, columnMap);
            case "screen":
                return createScreenData(row, columnMap);
            case "line item":
                return createLineItemData(row, columnMap);
            case "creative file":
                return createCreativeFileData(row, columnMap);
            case "dma":
                return createDmaData(row, columnMap);
            default:
                return null;
        }
    }

    private CampaignData createCampaignData(Row row, Map<String, Integer> columnMap) {
        return new CampaignData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("Campaign")),
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    private DateData createDateData(Row row, Map<String, Integer> columnMap) {
        String dateString = getCellValueAsString(row, columnMap.get("Date"));
        LocalDate date = null;
        if (dateString != null && !dateString.trim().isEmpty()) {
            try {
                date = LocalDate.parse(dateString, DATE_FORMATTER);
            } catch (Exception e) {
                System.err.println("Error parsing date: " + dateString);
            }
        }
        
        return new DateData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            date,
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    private PublisherData createPublisherData(Row row, Map<String, Integer> columnMap) {
        return new PublisherData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("Publisher ID")), // Additional ID field
            getCellValueAsString(row, columnMap.get("Publisher")),
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    private CategoryData createCategoryData(Row row, Map<String, Integer> columnMap) {
        String categoryName = getCellValueAsString(row, columnMap.get("Category name"));
        
        CategoryData data = new CategoryData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("Category group")),
            categoryName,
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
        
        return data;
    }

    private CityData createCityData(Row row, Map<String, Integer> columnMap) {
        return new CityData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("City")),
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    private ScreenData createScreenData(Row row, Map<String, Integer> columnMap) {
        return new ScreenData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("Screen ID")), // Additional ID field
            getCellValueAsString(row, columnMap.get("Screen")),
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Latitude")),
            getCellValueAsBigDecimal(row, columnMap.get("Longitude")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    private LineItemData createLineItemData(Row row, Map<String, Integer> columnMap) {
        return new LineItemData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("Line Item ID")), // Additional ID field
            getCellValueAsString(row, columnMap.get("Line item")),
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    private CreativeFileData createCreativeFileData(Row row, Map<String, Integer> columnMap) {
        return new CreativeFileData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("Creative File ID")), // Additional ID field
            getCellValueAsString(row, columnMap.get("Creative File")),
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    private DmaData createDmaData(Row row, Map<String, Integer> columnMap) {
        return new DmaData(
            getCellValueAsString(row, columnMap.get("Campaign ID")),
            getCellValueAsString(row, columnMap.get("DMA")),
            getCellValueAsLong(row, columnMap.get("Impressions")),
            getCellValueAsLong(row, columnMap.get("Playouts")),
            getCellValueAsDouble(row, columnMap.get("Imp. / Playout")),
            getCellValueAsBigDecimal(row, columnMap.get("Media CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Total CPM")),
            getCellValueAsBigDecimal(row, columnMap.get("Media Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Data Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Platform Costs")),
            getCellValueAsBigDecimal(row, columnMap.get("Invoice Amount")),
            getCellValueAsBigDecimal(row, columnMap.get("Client Margin")),
            getCellValueAsBigDecimal(row, columnMap.get("Total Spend"))
        );
    }

    // --- Helper methods for safe cell value extraction ---
    private String getCellValueAsString(Row row, Integer columnIndex) {
        if (columnIndex == null) return null;
        Cell cell = row.getCell(columnIndex);
        return getStringCellValue(cell);
    }

    private Double getCellValueAsDouble(Row row, Integer columnIndex) {
        if (columnIndex == null) return null;
        Cell cell = row.getCell(columnIndex);
        return getDoubleCellValue(cell);
    }

    private Long getCellValueAsLong(Row row, Integer columnIndex) {
        if (columnIndex == null) return null;
        Cell cell = row.getCell(columnIndex);
        return getLongCellValue(cell);
    }

    private BigDecimal getCellValueAsBigDecimal(Row row, Integer columnIndex) {
        if (columnIndex == null) return null;
        Cell cell = row.getCell(columnIndex);
        return getBigDecimalCellValue(cell);
    }

    // --- Original helper methods ---
    private static String getStringCellValue(Cell cell) {
        if (cell == null) return null;
        try {
            cell.setCellType(CellType.STRING);
            String value = cell.getStringCellValue();
            return (value == null || value.trim().isEmpty()) ? null : value.trim();
        } catch (Exception e) {
            return null;
        }
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
    
    // Method to clear all data collections
    private void clearAllData() {
        campaignDataList.clear();
        dateDataList.clear();
        publisherDataList.clear();
        categoryDataList.clear();
        cityDataList.clear();
        screenDataList.clear();
        lineItemDataList.clear();
        creativeFileDataList.clear();
        dmaDataList.clear();
    }
    
    // Method to process and store data in appropriate typed lists
    private void processAndStoreData(Sheet sheet, int headerRowIndex, Map<String, Integer> columnMap, String baseName) {
        
        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            
            // Check if row has any data
            if (isEmptyRow(row)) {
                continue;
            }
            
            
            try {
                Object dataObject = createDataObject(row, columnMap, baseName);
                if (dataObject != null) {
                    // Store in appropriate typed list based on sheet type
                    storeDataObject(dataObject, baseName);
                }
            } catch (Exception e) {
                System.err.println("Error processing row " + rowIndex + ": " + e.getMessage());
            }
        }
        
        // Print last 2 records from the appropriate list
        printLastTwoRecords(baseName);
    }
    
    // Method to store data object in appropriate typed list
    private void storeDataObject(Object dataObject, String baseName) {
        switch (baseName) {
            case "campaign":
                campaignDataList.add((CampaignData) dataObject);
                break;
            case "date":
                dateDataList.add((DateData) dataObject);
                break;
            case "publisher":
                publisherDataList.add((PublisherData) dataObject);
                break;
            case "category":
                categoryDataList.add((CategoryData) dataObject);
                break;
            case "city":
                cityDataList.add((CityData) dataObject);
                break;
            case "screen":
                screenDataList.add((ScreenData) dataObject);
                break;
            case "line item":
                lineItemDataList.add((LineItemData) dataObject);
                break;
            case "creative file":
                creativeFileDataList.add((CreativeFileData) dataObject);
                break;
            case "dma":
                dmaDataList.add((DmaData) dataObject);
                break;
        }
    }
    
    // Method to print last 2 records from appropriate list
    private void printLastTwoRecords(String baseName) {
        List<?> dataList = getDataList(baseName);
        if (dataList != null && !dataList.isEmpty()) {
            System.out.println("\n=== LAST TWO RECORDS FOR " + baseName.toUpperCase() + " ===");
            dataList.stream()
                .skip(Math.max(0, dataList.size() - 2))
                .forEach(System.out::println);
            System.out.println("Total " + baseName + " records: " + dataList.size());
        }
    }
    
    // Method to get appropriate data list
    private List<?> getDataList(String baseName) {
        switch (baseName) {
            case "campaign": return campaignDataList;
            case "date": return dateDataList;
            case "publisher": return publisherDataList;
            case "category": return categoryDataList;
            case "city": return cityDataList;
            case "screen": return screenDataList;
            case "line item": return lineItemDataList;
            case "creative file": return creativeFileDataList;
            case "dma": return dmaDataList;
            default: return null;
        }
    }
    
    // Method to generate charts with all collected data
    private void generateCharts() {
        System.out.println("\n=== GENERATING CHARTS ===");
        System.out.println("Campaign data: " + campaignDataList.size() + " records");
        System.out.println("Date data: " + dateDataList.size() + " records");
        System.out.println("Publisher data: " + publisherDataList.size() + " records");
        System.out.println("Category data: " + categoryDataList.size() + " records");
        System.out.println("City data: " + cityDataList.size() + " records");
        System.out.println("Screen data: " + screenDataList.size() + " records");
        System.out.println("Line item data: " + lineItemDataList.size() + " records");
        System.out.println("Creative file data: " + creativeFileDataList.size() + " records");
        System.out.println("DMA data: " + dmaDataList.size() + " records");
        
        // Generate PowerPoint with Campaign Overview
        try {
            String outputPath = "campaign_overview.pptx";
            pptService.createCampaignOverviewSlide(
                campaignDataList, dateDataList, cityDataList, 
                categoryDataList, screenDataList, publisherDataList, outputPath);
        } catch (Exception e) {
            System.err.println("Error generating PowerPoint: " + e.getMessage());
            e.printStackTrace();
        }
        
        System.out.println("Chart generation completed.");
    }
    
    // Getter methods to access the data collections
    public List<CampaignData> getCampaignDataList() { return campaignDataList; }
    public List<DateData> getDateDataList() { return dateDataList; }
    public List<PublisherData> getPublisherDataList() { return publisherDataList; }
    public List<CategoryData> getCategoryDataList() { return categoryDataList; }
    public List<CityData> getCityDataList() { return cityDataList; }
    public List<ScreenData> getScreenDataList() { return screenDataList; }
    public List<LineItemData> getLineItemDataList() { return lineItemDataList; }
    public List<CreativeFileData> getCreativeFileDataList() { return creativeFileDataList; }
    public List<DmaData> getDmaDataList() { return dmaDataList; }
}