package com.example.excelchart.controller;

import com.example.excelchart.service.ExcelService;
import com.example.excelchart.service.PowerPointExportService;
import com.example.excelchart.service.CampaignOverviewPPTService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.UUID;

@Controller
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @Autowired
    private PowerPointExportService pptService;
    
    @Autowired
    private CampaignOverviewPPTService campaignOverviewPPTService;
    
    @Value("${app.upload.dir:${java.io.tmpdir}/uploads}")
    private String uploadDir;

    @PostMapping("/upload")
    public String uploadExcel(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) {
        try {
            // Validate file
            if (file.isEmpty()) {
                redirectAttributes.addFlashAttribute("error", "Please select a file to upload.");
                return "redirect:/index.html";
            }
            
            // Check file extension
            String originalFilename = file.getOriginalFilename();
            if (originalFilename == null || !originalFilename.toLowerCase().endsWith(".xlsx")) {
                redirectAttributes.addFlashAttribute("error", "Please upload an Excel (.xlsx) file.");
                return "redirect:/index.html";
            }
            
            System.out.println("Processing file: " + originalFilename + " (Size: " + file.getSize() + " bytes)");
            
            // Create upload directory
            Path uploadPath = Path.of(uploadDir);
            Files.createDirectories(uploadPath);
            
            // Generate unique filename to avoid conflicts
            String uniqueFilename = UUID.randomUUID().toString() + "_" + originalFilename;
            File excelFile = uploadPath.resolve(uniqueFilename).toFile();
            
            // Save the uploaded file
            file.transferTo(excelFile);
            System.out.println("File saved to: " + excelFile.getAbsolutePath());
            
            // Process the Excel file
            System.out.println("Starting Excel processing...");
            excelService.readAndGenerateCharts(excelFile);
            System.out.println("Excel processing completed successfully!");
            
            // Clean up the uploaded file
            Files.deleteIfExists(excelFile.toPath());
            System.out.println("Temporary file cleaned up.");
            
            // Generate PowerPoint report
            try {
                String outputPath = "campaign_overview.pptx";
                campaignOverviewPPTService.createCampaignOverviewSlide(
                    excelService.getCampaignDataList(),
                    excelService.getDateDataList(),
                    excelService.getCityDataList(),
                    excelService.getCategoryDataList(),
                    excelService.getScreenDataList(),
                    excelService.getPublisherDataList(),
                    outputPath
                );
                System.out.println("PowerPoint report generated successfully!");
                redirectAttributes.addFlashAttribute("success", "Excel file processed and PowerPoint report generated successfully!");
                return "redirect:/index.html?reportGenerated=true";
            } catch (Exception e) {
                System.err.println("Error generating PowerPoint: " + e.getMessage());
                redirectAttributes.addFlashAttribute("success", "Excel file processed successfully, but PowerPoint generation failed.");
                return "redirect:/index.html?reportGenerated=false";
            }
            
        } catch (Exception e) {
            System.err.println("Error processing file: " + e.getMessage());
            e.printStackTrace();
            redirectAttributes.addFlashAttribute("error", "Error processing file: " + e.getMessage());
            return "redirect:/index.html";
        }
    }

    @GetMapping("/test-excel")
    @ResponseBody
    public String testExcelReading() {
        try {
            // Read the Excel file from resources/data folder
            File excelFile = new File("src/main/resources/data/Ethos - Dog - Report (3).xlsx");
            if (!excelFile.exists()) {
                return "Excel file not found at: " + excelFile.getAbsolutePath();
            }
            
            excelService.readAndGenerateCharts(excelFile);
            return "Excel file processed successfully! Check console for details.";
            
        } catch (Exception e) {
            return "Error processing Excel file: " + e.getMessage();
        }
    }

    @GetMapping("/download-report")
    @ResponseBody
    public ResponseEntity<Resource> downloadReport() {
        try {
            File reportFile = new File("campaign_overview.pptx");
            if (!reportFile.exists()) {
                return ResponseEntity.notFound().build();
            }
            
            Path path = reportFile.toPath();
            Resource resource = new org.springframework.core.io.FileSystemResource(path);
            
            return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"campaign_overview.pptx\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
                
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    @GetMapping("/")
    public String index() {
        return "redirect:/index.html";
    }
}