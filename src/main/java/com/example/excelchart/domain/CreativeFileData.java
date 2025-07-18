package com.example.excelchart.domain;

import java.math.BigDecimal;

public class CreativeFileData {
    private String campaignId;
    private String creativeFileId;
    private String creativeFile;
    private Long impressions;
    private Long playouts;
    private Double impPerPlayout;
    private BigDecimal mediaCpm;
    private BigDecimal totalCpm;
    private BigDecimal mediaCosts;
    private BigDecimal dataCosts;
    private BigDecimal platformCosts;
    private BigDecimal invoiceAmount;
    private BigDecimal clientMargin;
    private BigDecimal totalSpend;

    public CreativeFileData(String campaignId, String creativeFileId, String creativeFile, Long impressions, Long playouts,
                            Double impPerPlayout, BigDecimal mediaCpm, BigDecimal totalCpm,
                            BigDecimal mediaCosts, BigDecimal dataCosts, BigDecimal platformCosts,
                            BigDecimal invoiceAmount, BigDecimal clientMargin, BigDecimal totalSpend) {
        this.campaignId = campaignId;
        this.creativeFileId = creativeFileId;
        this.creativeFile = creativeFile;
        this.impressions = impressions;
        this.playouts = playouts;
        this.impPerPlayout = impPerPlayout;
        this.mediaCpm = mediaCpm;
        this.totalCpm = totalCpm;
        this.mediaCosts = mediaCosts;
        this.dataCosts = dataCosts;
        this.platformCosts = platformCosts;
        this.invoiceAmount = invoiceAmount;
        this.clientMargin = clientMargin;
        this.totalSpend = totalSpend;
    }

    // Getters
    public String getCampaignId() { return campaignId; }
    public String getCreativeFileId() { return creativeFileId; }
    public String getCreativeFile() { return creativeFile; }
    public Long getImpressions() { return impressions; }
    public Long getPlayouts() { return playouts; }
    public Double getImpPerPlayout() { return impPerPlayout; }
    public BigDecimal getMediaCpm() { return mediaCpm; }
    public BigDecimal getTotalCpm() { return totalCpm; }
    public BigDecimal getMediaCosts() { return mediaCosts; }
    public BigDecimal getDataCosts() { return dataCosts; }
    public BigDecimal getPlatformCosts() { return platformCosts; }
    public BigDecimal getInvoiceAmount() { return invoiceAmount; }
    public BigDecimal getClientMargin() { return clientMargin; }
    public BigDecimal getTotalSpend() { return totalSpend; }

    @Override
    public String toString() {
        return "CreativeFileData{" +
                "campaignId='" + campaignId + '\'' +
                ", creativeFileId='" + creativeFileId + '\'' +
                ", creativeFile='" + creativeFile + '\'' +
                ", impressions=" + impressions +
                ", playouts=" + playouts +
                ", impPerPlayout=" + impPerPlayout +
                ", mediaCpm=" + mediaCpm +
                ", totalCpm=" + totalCpm +
                ", mediaCosts=" + mediaCosts +
                ", dataCosts=" + dataCosts +
                ", platformCosts=" + platformCosts +
                ", invoiceAmount=" + invoiceAmount +
                ", clientMargin=" + clientMargin +
                ", totalSpend=" + totalSpend +
                '}';
    }
}