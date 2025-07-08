package com.example.excelchart.domain;


import java.math.BigDecimal;

public class CityData {
    private String campaignId;
    private String city;
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

    public CityData(String campaignId, String city, Long impressions, Long playouts,
                    Double impPerPlayout, BigDecimal mediaCpm, BigDecimal totalCpm,
                    BigDecimal mediaCosts, BigDecimal dataCosts, BigDecimal platformCosts,
                    BigDecimal invoiceAmount, BigDecimal clientMargin, BigDecimal totalSpend) {
        this.campaignId = campaignId;
        this.city = city;
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
    public String getCity() { return city; }
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
        return "CityData{" +
                "campaignId='" + campaignId + '\'' +
                ", city='" + city + '\'' +
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