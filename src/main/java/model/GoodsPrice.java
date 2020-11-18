package model;

import java.math.BigDecimal;

public class GoodsPrice {
    private BigDecimal price;
    private BigDecimal responsePrice;
    private String goodsCode;
    private String goodsId;
    private String goodsName;
    private String bidderId;
    private String bidderName;
    private String companyId;
    private String companyName;

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public BigDecimal getResponsePrice() {
        return responsePrice;
    }

    public void setResponsePrice(BigDecimal responsePrice) {
        this.responsePrice = responsePrice;
    }

    public String getGoodsCode() {
        return goodsCode;
    }

    public void setGoodsCode(String goodsCode) {
        this.goodsCode = goodsCode;
    }

    public String getGoodsId() {
        return goodsId;
    }

    public void setGoodsId(String goodsId) {
        this.goodsId = goodsId;
    }

    public String getGoodsName() {
        return goodsName;
    }

    public void setGoodsName(String goodsName) {
        this.goodsName = goodsName;
    }

    public String getBidderId() {
        return bidderId;
    }

    public void setBidderId(String bidderId) {
        this.bidderId = bidderId;
    }

    public String getBidderName() {
        return bidderName;
    }

    public void setBidderName(String bidderName) {
        this.bidderName = bidderName;
    }

    public String getCompanyId() {
        return companyId;
    }

    public void setCompanyId(String companyId) {
        this.companyId = companyId;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    @Override
    public String toString() {
        return "GoodsPrice{" +
                "price=" + price +
                ", responsePrice=" + responsePrice +
                ", goodsCode='" + goodsCode + '\'' +
                ", goodsId='" + goodsId + '\'' +
                ", goodsName='" + goodsName + '\'' +
                ", bidderId='" + bidderId + '\'' +
                ", bidderName='" + bidderName + '\'' +
                ", companyId='" + companyId + '\'' +
                ", companyName='" + companyName + '\'' +
                '}';
    }
}
