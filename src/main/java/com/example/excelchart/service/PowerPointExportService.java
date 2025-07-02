package com.example.excelchart.service;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.*;
import java.util.*;

@Service
public class PowerPointExportService {

    public void exportChartsToPPT(List<String> chartPaths, String pptFilePath) throws Exception {
        XMLSlideShow ppt = new XMLSlideShow();

        for (String chartPath : chartPaths) {
            byte[] pictureData = Files.readAllBytes(Paths.get(chartPath));
            XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);
            XSLFSlide slide = ppt.createSlide();
            XSLFPictureShape pic = slide.createPicture(pd);
            pic.setAnchor(new java.awt.Rectangle(50, 50, 640, 480));
        }

        try (FileOutputStream out = new FileOutputStream(pptFilePath)) {
            ppt.write(out);
        }
    }
}
