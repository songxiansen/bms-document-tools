package bc.bms.common.workbook.model;

import java.math.BigDecimal;

public class ItemData extends Item {

    private int dataType;

    private String dateFlag;

    private BigDecimal value;

    public ItemData() {
        super();
    }

    public ItemData(Integer id, String name) {
        super(id, name);
    }

    public ItemData(Integer id, String name, boolean editable) {
        super(id, name, editable);
    }

    public ItemData(Integer id, String name, boolean editable, int dataType, String dateFlag, BigDecimal value) {
        super(id, name, editable);

        this.dataType = dataType;
        this.dateFlag = dateFlag;
        this.value = value;
    }

    public int getDataType() {
        return dataType;
    }

    public void setDataType(int dataType) {
        this.dataType = dataType;
    }

    public String getDateFlag() {
        return dateFlag;
    }

    public void setDateFlag(String dateFlag) {
        this.dateFlag = dateFlag;
    }

    public BigDecimal getValue() {
        return value;
    }

    public void setValue(BigDecimal value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return "ItemData{" +
                "dataType=" + dataType +
                ", dateFlag='" + dateFlag + '\'' +
                ", value=" + value +
                '}';
    }

}
