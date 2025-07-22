package com.example.excelchart.service;

import com.example.excelchart.domain.CategoryData;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.springframework.stereotype.Service;
import java.util.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

@Service
public class ExcelBarChartGeneratorService {
    public void addCategoryBarChart(List<CategoryData> categoryDataList, XSSFWorkbook workbook, Sheet chartSheet) {
        if (categoryDataList == null || categoryDataList.isEmpty()) {
            System.out.println("No category data for bar chart.");
            return;
        }
        // Aggregate total spend by category name (after last '>')
        Map<String, Double> categorySpend = new LinkedHashMap<>();
        for (CategoryData data : categoryDataList) {
            String rawName = data.getCategoryName();
            if (rawName == null || rawName.trim().isEmpty()) continue;
            String categoryName = rawName.contains(">") ? rawName.substring(rawName.lastIndexOf(">") + 1).trim() : rawName.trim();
            double spend = data.getTotalSpend() != null ? data.getTotalSpend().doubleValue() : 0.0;
            categorySpend.put(categoryName, categorySpend.getOrDefault(categoryName, 0.0) + spend);
        }
        if (categorySpend.isEmpty()) {
            System.out.println("No valid category names for bar chart.");
            return;
        }
        // Write data to a new sheet (or reuse data sheet if desired)
        XSSFSheet dataSheet = (XSSFSheet) workbook.getSheet("Bar Data");
        if (dataSheet == null) {
            dataSheet = workbook.createSheet("Bar Data");
        }
        // Write header
        org.apache.poi.ss.usermodel.Row header = dataSheet.createRow(0);
        header.createCell(0).setCellValue("Category");
        header.createCell(1).setCellValue("Total Spend");
        // Write data
        int rowIdx = 1;
        for (Map.Entry<String, Double> entry : categorySpend.entrySet()) {
            org.apache.poi.ss.usermodel.Row row = dataSheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getValue());
        }
        dataSheet.autoSizeColumn(0);
        dataSheet.autoSizeColumn(1);
        System.out.println("Bar chart data rows: " + (rowIdx - 1));
        // Create the bar chart in its own sheet
        XSSFSheet barChartSheet = (XSSFSheet) workbook.getSheet("Bar Chart");
        if (barChartSheet == null) {
            barChartSheet = workbook.createSheet("Bar Chart");
        }
        Drawing<?> drawing = barChartSheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 10, 20);
        System.out.println("Bar chart anchor: columns 1-10, rows 1-20");
        XSSFChart chart = ((XSSFDrawing) drawing).createChart(anchor);
        chart.setTitleText("Total Spend by Category");
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);
        int lastRow = rowIdx - 1;
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
            dataSheet, new CellRangeAddress(1, lastRow, 0, 0));
        XDDFNumericalDataSource<Double> spends = XDDFDataSourcesFactory.fromNumericCellRange(
            dataSheet, new CellRangeAddress(1, lastRow, 1, 1));
        XDDFChartData data = chart.createData(ChartTypes.BAR, null, null);
        XDDFChartData.Series series = data.addSeries(categories, spends);
        series.setTitle("Total Spend", null);
        chart.plot(data);
        // Set bar chart to be vertical (column chart)
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
    }
} 