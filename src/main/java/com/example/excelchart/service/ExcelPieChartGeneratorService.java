package com.example.excelchart.service;

import com.example.excelchart.domain.CategoryData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos;
import org.springframework.stereotype.Service;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.FileOutputStream;
import java.util.*;

@Service
public class ExcelPieChartGeneratorService {
    @Autowired
    private ExcelBarChartGeneratorService excelBarChartGeneratorService;

    public void generateLocationPieChartExcel(List<CategoryData> categoryDataList) throws Exception {
        if (categoryDataList == null || categoryDataList.isEmpty()) {
            System.out.println("No location data for Excel pie chart.");
            return;
        }

        // Aggregate impressions by location name
        Map<String, Long> locationImpressions = new LinkedHashMap<>();
        for (CategoryData data : categoryDataList) {
            String rawName = data.getCategoryName();
            if (rawName == null || rawName.trim().isEmpty()) continue;
            String locationName = rawName.contains(">") ? rawName.substring(rawName.lastIndexOf(">") + 1).trim() : rawName.trim();
            long impressions = data.getImpressions() != null ? data.getImpressions() : 0L;
            locationImpressions.put(locationName, locationImpressions.getOrDefault(locationName, 0L) + impressions);
        }

        if (locationImpressions.isEmpty()) {
            System.out.println("No valid location names for Excel pie chart.");
            return;
        }

        // Create workbook and sheets
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet dataSheet = workbook.createSheet("Location Data");
        Sheet chartSheet = workbook.createSheet("Pie Chart");

        // Write header
        Row header = dataSheet.createRow(0);
        header.createCell(0).setCellValue("Location");
        header.createCell(1).setCellValue("Impressions");

        // Write data
        int rowIdx = 1;
        for (Map.Entry<String, Long> entry : locationImpressions.entrySet()) {
            Row row = dataSheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getValue());
        }

        // Autosize columns
        dataSheet.autoSizeColumn(0);
        dataSheet.autoSizeColumn(1);

        // Create the pie chart
        Drawing<?> drawing = chartSheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 10, 20);
        XSSFChart chart = ((XSSFDrawing) drawing).createChart(anchor);
        chart.setTitleText("Impressions by Location");
        chart.setTitleOverlay(false);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        // Data sources
        int lastRow = locationImpressions.size();
        XDDFDataSource<String> locations = XDDFDataSourcesFactory.fromStringCellRange(
                dataSheet, new CellRangeAddress(1, lastRow, 0, 0));
        XDDFNumericalDataSource<Double> impressions = XDDFDataSourcesFactory.fromNumericCellRange(
                dataSheet, new CellRangeAddress(1, lastRow, 1, 1));

        // Pie chart
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        data.addSeries(locations, impressions);
        chart.plot(data);

        // Enable labels: category name + percentage, hide value
        if (data instanceof XDDFPieChartData) {
            XDDFPieChartData pieData = (XDDFPieChartData) data;
            if (pieData.getSeriesCount() > 0) {
                XDDFChartData.Series series = pieData.getSeries(0);
                try {
                    java.lang.reflect.Method getXmlObject = series.getClass().getMethod("getXmlObject");
                    CTPieSer ctSeries = (CTPieSer) getXmlObject.invoke(series);
                    CTDLbls dLbls = ctSeries.getDLbls();
                    if (dLbls == null) {
                        dLbls = ctSeries.addNewDLbls();
                    }
                    dLbls.addNewShowCatName().setVal(true);      // ✅ Show category name
                    dLbls.addNewShowPercent().setVal(true);      // ✅ Show percentage
                    dLbls.addNewShowVal().setVal(false);         // ❌ Hide raw value
                    dLbls.addNewShowLegendKey().setVal(false);   // Optional
                    dLbls.addNewShowLeaderLines().setVal(true);  // ✅ Show connector lines
                    dLbls.addNewDLblPos().setVal(STDLblPos.OUT_END); // ✅ Position labels outside
                } catch (Exception e) {
                    System.err.println("Could not set pie chart data labels: " + e.getMessage());
                }
            }
        }

        // Save workbook
        // Add bar chart next to pie chart
        excelBarChartGeneratorService.addCategoryBarChart(categoryDataList, workbook, chartSheet);
        try (FileOutputStream fileOut = new FileOutputStream("location_pie_chart.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();

        System.out.println("Excel file with pie chart saved as location_pie_chart.xlsx");
    }
}
