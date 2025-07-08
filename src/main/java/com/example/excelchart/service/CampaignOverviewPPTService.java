package com.example.excelchart.service;

import com.example.excelchart.domain.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.sl.usermodel.Placeholder;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class CampaignOverviewPPTService {

    public void createCampaignOverviewSlide(List<CampaignData> campaignDataList,
                                          List<DateData> dateDataList,
                                          List<CityData> cityDataList,
                                          List<CategoryData> categoryDataList,
                                          List<ScreenData> screenDataList,
                                          List<PublisherData> publisherDataList,
                                          String outputPath) throws IOException {
        // Create a new PowerPoint presentation
        XMLSlideShow ppt = new XMLSlideShow();
        // Set slide size to standard 16:9
        ppt.setPageSize(new Dimension(1280, 720));
        // Create the Campaign Overview slide
        XSLFSlide slide = ppt.createSlide();
        // Add title
        addTitle(slide, "Campaign Overview");
        // Calculate and add campaign details
        addCampaignDetails(slide, campaignDataList, dateDataList, cityDataList, 
                          categoryDataList, screenDataList, publisherDataList);
        // Save the presentation
        try (FileOutputStream out = new FileOutputStream(outputPath)) {
            ppt.write(out);
        }
        System.out.println("Campaign Overview PowerPoint created at: " + outputPath);
    }
    
    private void addTitle(XSLFSlide slide, String title) {
        XSLFTextBox titleBox = slide.createTextBox();
        titleBox.setAnchor(new Rectangle(50, 30, 1180, 60));
        
        XSLFTextParagraph titlePara = titleBox.addNewTextParagraph();
        titlePara.setTextAlign(TextAlign.CENTER);
        
        XSLFTextRun titleRun = titlePara.addNewTextRun();
        titleRun.setText(title);
        titleRun.setFontSize(32.0);
        titleRun.setBold(true);
        titleRun.setFontColor(Color.BLACK);
    }
    
    private void addCampaignDetails(XSLFSlide slide, 
                                  List<CampaignData> campaignDataList,
                                  List<DateData> dateDataList,
                                  List<CityData> cityDataList,
                                  List<CategoryData> categoryDataList,
                                  List<ScreenData> screenDataList,
                                  List<PublisherData> publisherDataList) {
        
        // Calculate campaign details
        CampaignOverviewData overviewData = calculateOverviewData(
            campaignDataList, dateDataList, cityDataList, categoryDataList, 
            screenDataList, publisherDataList);
        
        // Create content area
        XSLFTextBox contentBox = slide.createTextBox();
        contentBox.setAnchor(new Rectangle(50, 120, 1180, 570));
        
        // Add each detail row
        addDetailRow(contentBox, "Flight", overviewData.getFlight(), 0);
        addDetailRow(contentBox, "Budget", "N/A", 1);
        addDetailRow(contentBox, "Geo", overviewData.getGeo(), 2);
        addDetailRow(contentBox, "Placements", overviewData.getPlacements(), 3);
        addDetailRow(contentBox, "Screens", overviewData.getScreens(), 4);
        addDetailRow(contentBox, "Publisher", overviewData.getPublisher(), 5);
        addDetailRow(contentBox, "Impressions Served", overviewData.getImpressionsServed(), 6);
        addDetailRow(contentBox, "eCPM", overviewData.getEcpm(), 7);
        addDetailRow(contentBox, "Spend", overviewData.getSpend(), 8);
    }
    
    private void addDetailRow(XSLFTextBox contentBox, String label, String value, int rowIndex) {
        int yPosition = 20 + (rowIndex * 60);
        
        // Add label
        XSLFTextParagraph labelPara = contentBox.addNewTextParagraph();
        labelPara.setLeftMargin(0.0);
        labelPara.setIndentLevel(0);
        
        XSLFTextRun labelRun = labelPara.addNewTextRun();
        labelRun.setText(label + ": ");
        labelRun.setFontSize(18.0);
        labelRun.setBold(true);
        labelRun.setFontColor(new Color(68, 84, 106));
        
        // Add value
        XSLFTextParagraph valuePara = contentBox.addNewTextParagraph();
        valuePara.setLeftMargin(200.0);
        valuePara.setIndentLevel(0);
        
        XSLFTextRun valueRun = valuePara.addNewTextRun();
        valueRun.setText(value);
        valueRun.setFontSize(18.0);
        valueRun.setFontColor(Color.BLACK);
    }
    
    private CampaignOverviewData calculateOverviewData(List<CampaignData> campaignDataList,
                                                     List<DateData> dateDataList,
                                                     List<CityData> cityDataList,
                                                     List<CategoryData> categoryDataList,
                                                     List<ScreenData> screenDataList,
                                                     List<PublisherData> publisherDataList) {
        
        CampaignOverviewData data = new CampaignOverviewData();
        
        // Flight: Find earliest and latest dates
        if (!dateDataList.isEmpty()) {
            LocalDate earliestDate = dateDataList.stream()
                .map(DateData::getDate)
                .filter(date -> date != null)
                .min(LocalDate::compareTo)
                .orElse(null);
                
            LocalDate latestDate = dateDataList.stream()
                .map(DateData::getDate)
                .filter(date -> date != null)
                .max(LocalDate::compareTo)
                .orElse(null);
                
            if (earliestDate != null && latestDate != null) {
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMM dd");
                data.setFlight(earliestDate.format(formatter) + " - " + latestDate.format(formatter));
            } else {
                data.setFlight("N/A");
            }
        } else {
            data.setFlight("N/A");
        }
        
        // Geo: Get all unique cities
        if (!cityDataList.isEmpty()) {
            String cities = cityDataList.stream()
                .map(CityData::getCity)
                .filter(city -> city != null && !city.trim().isEmpty())
                .distinct()
                .collect(Collectors.joining(", "));
            data.setGeo(cities.isEmpty() ? "N/A" : cities);
        } else {
            data.setGeo("N/A");
        }
        
        // Placements: Get category names excluding "Outdoor"
        System.out.println("Processing placements from " + categoryDataList.size() + " category records");
        if (!categoryDataList.isEmpty()) {
            // Debug: Print all category names
            System.out.println("All category names found:");
            categoryDataList.forEach(cat -> {
                System.out.println("  - Category: '" + cat.getCategoryName() + "'");
            });
            
            String placements = categoryDataList.stream()
            .map(CategoryData::getCategoryName)
            .filter(category -> category != null && !category.trim().isEmpty())
            .map(category -> category.contains(">") ? category.split(">")[1].trim() : category.trim())
            .distinct()
            .collect(Collectors.joining(", "));
        
            
            System.out.println("Final placements string: '" + placements + "'");
            data.setPlacements(placements.isEmpty() ? "N/A" : placements);
        } else {
            System.out.println("No category data found - setting placements to N/A");
            data.setPlacements("N/A");
        }
        
        // Screens: Number of records - 1
        data.setScreens(String.valueOf(Math.max(0, screenDataList.size() - 1)));
        
        // Publisher: Get all unique publishers
        if (!publisherDataList.isEmpty()) {
            String publishers = publisherDataList.stream()
                .map(PublisherData::getPublisher)
                .filter(pub -> pub != null && !pub.trim().isEmpty())
                .distinct()
                .collect(Collectors.joining(", "));
            data.setPublisher(publishers.isEmpty() ? "N/A" : publishers);
        } else {
            data.setPublisher("N/A");
        }
        
        // Impressions Served, eCPM, Spend: Get from last row of campaign data
        if (!campaignDataList.isEmpty()) {
            CampaignData lastCampaign = campaignDataList.get(campaignDataList.size()-1);
            
            // Impressions Served (show full number)
            if (lastCampaign.getImpressions() != null) {
                data.setImpressionsServed(String.valueOf(lastCampaign.getImpressions()));
            } else {
                data.setImpressionsServed("N/A");
            }
            
            // eCPM (Total CPM, 2 decimal places)
            if (lastCampaign.getTotalCpm() != null) {
                data.setEcpm("$" + String.format("%.2f", lastCampaign.getTotalCpm()));
            } else {
                data.setEcpm("N/A");
            }
            
            // Spend (Total Spend, 2 decimal places)
            if (lastCampaign.getTotalSpend() != null) {
                data.setSpend("$" + String.format("%.2f", lastCampaign.getTotalSpend()));
            } else {
                data.setSpend("N/A");
            }
        } else {
            data.setImpressionsServed("N/A");
            data.setEcpm("N/A");
            data.setSpend("N/A");
        }
        
        return data;
    }
    
    // Data class to hold campaign overview information
    private static class CampaignOverviewData {
        private String flight;
        private String geo;
        private String placements;
        private String screens;
        private String publisher;
        private String impressionsServed;
        private String ecpm;
        private String spend;
        
        // Getters and setters
        public String getFlight() { return flight; }
        public void setFlight(String flight) { this.flight = flight; }
        
        public String getGeo() { return geo; }
        public void setGeo(String geo) { this.geo = geo; }
        
        public String getPlacements() { return placements; }
        public void setPlacements(String placements) { this.placements = placements; }
        
        public String getScreens() { return screens; }
        public void setScreens(String screens) { this.screens = screens; }
        
        public String getPublisher() { return publisher; }
        public void setPublisher(String publisher) { this.publisher = publisher; }
        
        public String getImpressionsServed() { return impressionsServed; }
        public void setImpressionsServed(String impressionsServed) { this.impressionsServed = impressionsServed; }
        
        public String getEcpm() { return ecpm; }
        public void setEcpm(String ecpm) { this.ecpm = ecpm; }
        
        public String getSpend() { return spend; }
        public void setSpend(String spend) { this.spend = spend; }
    }
} 