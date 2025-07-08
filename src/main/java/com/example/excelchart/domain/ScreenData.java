package com.example.excelchart.domain;


import java.math.BigDecimal;

public class ScreenData {
    private String campaignId;
    private String screenId;
    private String screen;
    private Long impressions;
    private Long playouts;
    private Double impPerPlayout;
    private BigDecimal mediaCpm;
    private BigDecimal totalCpm;
    private BigDecimal mediaCosts;
    private BigDecimal latitude;
    private BigDecimal longitude;
    private BigDecimal dataCosts;
    private BigDecimal platformCosts;
    private BigDecimal invoiceAmount;
    private BigDecimal clientMargin;
    private BigDecimal totalSpend;

    public ScreenData(String campaignId, String screenId, String screen, Long impressions, Long playouts,
                      Double impPerPlayout, BigDecimal mediaCpm, BigDecimal totalCpm,
                      BigDecimal mediaCosts, BigDecimal latitude, BigDecimal longitude,
                      BigDecimal dataCosts, BigDecimal platformCosts,
                      BigDecimal invoiceAmount, BigDecimal clientMargin, BigDecimal totalSpend) {
        this.campaignId = campaignId;
        this.screenId = screenId;
        this.screen = screen;
        this.impressions = impressions;
        this.playouts = playouts;
        this.impPerPlayout = impPerPlayout;
        this.mediaCpm = mediaCpm;
        this.totalCpm = totalCpm;
        this.mediaCosts = mediaCosts;
        this.latitude = latitude;
        this.longitude = longitude;
        this.dataCosts = dataCosts;
        this.platformCosts = platformCosts;
        this.invoiceAmount = invoiceAmount;
        this.clientMargin = clientMargin;
        this.totalSpend = totalSpend;
    }

    // Getters
    public String getCampaignId() { return campaignId; }
    public String getScreenId() { return screenId; }
    public String getScreen() { return screen; }
    public Long getImpressions() { return impressions; }
    public Long getPlayouts() { return playouts; }
    public Double getImpPerPlayout() { return impPerPlayout; }
    public BigDecimal getMediaCpm() { return mediaCpm; }
    public BigDecimal getTotalCpm() { return totalCpm; }
    public BigDecimal getMediaCosts() { return mediaCosts; }
    public BigDecimal getLatitude() { return latitude; }
    public BigDecimal getLongitude() { return longitude; }
    public BigDecimal getDataCosts() { return dataCosts; }
    public BigDecimal getPlatformCosts() { return platformCosts; }
    public BigDecimal getInvoiceAmount() { return invoiceAmount; }
    public BigDecimal getClientMargin() { return clientMargin; }
    public BigDecimal getTotalSpend() { return totalSpend; }

    @Override
    public String toString() {
        return "ScreenData{" +
                "campaignId='" + campaignId + '\'' +
                ", screenId='" + screenId + '\'' +
                ", screen='" + screen + '\'' +
                ", impressions=" + impressions +
                ", playouts=" + playouts +
                ", impPerPlayout=" + impPerPlayout +
                ", mediaCpm=" + mediaCpm +
                ", totalCpm=" + totalCpm +
                ", mediaCosts=" + mediaCosts +
                ", latitude=" + latitude +
                ", longitude=" + longitude +
                ", dataCosts=" + dataCosts +
                ", platformCosts=" + platformCosts +
                ", invoiceAmount=" + invoiceAmount +
                ", clientMargin=" + clientMargin +
                ", totalSpend=" + totalSpend +
                '}';
    }
}