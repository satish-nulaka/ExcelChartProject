package com.example.excelchart.controller;

import com.example.excelchart.service.ExcelService;
import com.example.excelchart.service.PowerPointExportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

@Controller
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @Autowired
    private PowerPointExportService pptService;

@PostMapping("/upload")
public String uploadExcel(@RequestParam("file") MultipartFile file) throws Exception {
    // Use a directory inside the system temp directory for uploads
    Path uploadDir = Path.of(System.getProperty("java.io.tmpdir"), "uploads");
    Files.createDirectories(uploadDir);

    // Save the uploaded file to the uploads directory
    File excelFile = uploadDir.resolve(file.getOriginalFilename()).toFile();
    file.transferTo(excelFile);

    excelService.readAndGenerateCharts(excelFile);

    pptService.exportChartsToPPT(
        Arrays.asList("chart-Sheet1.png", "chart-Sheet2.png", "chart-Sheet3.png"),
        "output-presentation.pptx"
    );

    return "redirect:/report.html";
}

    @GetMapping("/")
    public String index() {
        return "redirect:/index.html";
    }
}