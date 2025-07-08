package com.example.excelchart;

import com.example.excelchart.service.ExcelService;
import java.io.File;

public class ExcelReaderTest {
    
    public static void main(String[] args) {
        try {
            // Set the path to your Excel file
            File excelFile = new File("src/main/resources/data/Ethos - Dog - Report (3).xlsx");
            
            if (!excelFile.exists()) {
                System.err.println("Excel file not found at: " + excelFile.getAbsolutePath());
                return;
            }
            
            System.out.println("Testing Excel reader with file: " + excelFile.getName());
            
            // Create ExcelService instance
            ExcelService excelService = new ExcelService();
            
            // Test the Excel reader
            excelService.readAndGenerateCharts(excelFile);
            
            System.out.println("Excel file processed successfully!");
            
        } catch (Exception e) {
            System.err.println("Error processing Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }
} 