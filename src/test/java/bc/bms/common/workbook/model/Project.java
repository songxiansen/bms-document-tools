package bc.bms.common.workbook.model;

import java.math.BigDecimal;

public class Project {

    private Integer id;

    private String code;

    private String name;

    private BigDecimal costAmount;

    private BigDecimal fundAmount;

    private String status;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public BigDecimal getCostAmount() {
        return costAmount;
    }

    public void setCostAmount(BigDecimal costAmount) {
        this.costAmount = costAmount;
    }

    public BigDecimal getFundAmount() {
        return fundAmount;
    }

    public void setFundAmount(BigDecimal fundAmount) {
        this.fundAmount = fundAmount;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public Project() {

    }

    public Project(Integer id, String code, String name, BigDecimal costAmount, BigDecimal fundAmount, String status) {
        this.id = id;
        this.code = code;
        this.name = name;
        this.costAmount = costAmount;
        this.fundAmount = fundAmount;
        this.status = status;
    }

    @Override
    public String toString() {
        return "Project{" +
                "id=" + id +
                ", code='" + code + '\'' +
                ", name='" + name + '\'' +
                ", costAmount=" + costAmount +
                ", fundAmount=" + fundAmount +
                ", status='" + status + '\'' +
                '}';
    }
}
