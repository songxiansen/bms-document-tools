package bc.bms.common.workbook.model;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.codehaus.jackson.annotate.JsonIgnore;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * 工作表列定义模型
 *
 * @param <T> 列模型
 */
public class ColumnDefinition<T> implements FormulaCell, Serializable {

    /**
     * 列的数据模型名称存储Field name
     */
    public static final String CLASS_NAME_OF_COLUMN_MODEL_FIELD_NAME = "classNameOfColumnModel";

    @JsonIgnore
    private String name;

    /**
     * 存储值的Field name
     */
    private String fieldNameOfValue;

    /**
     * 当列数据模型是Map是，存储value type name of Map
     */
    private String fieldTypeOfValue;

    /**
     * 值格式，目前仅支持DecimalFormat
     */
    private String valueFormatPattern;

    /**
     * 列数据示例，主要存储该列数据共有属性
     */
    private T sampleColumnModel;

    /**
     * 列数据模型名称，直接从@sampleColumnModel获得
     */
    private String classNameOfColumnModel;

    /**
     * 数据源列索引
     */
    private Integer dataColumnIndex;

    @JsonIgnore
    private String formula;

    /**
     * 获得公式定义的列索引，用于从其他列复制公式。
     */
    @JsonIgnore
    private Integer formulaColumnIndex;

    /**
     * 列是否可编辑
     */
    @JsonIgnore
    private Boolean editable;

    /**
     * 列中只读单元格背景颜色
     */
    @JsonIgnore
    private XSSFColor readOnlyCellColor;

    /**
     * 列中可编辑单元格背景颜色
     */
    @JsonIgnore
    private XSSFColor editableCellColor;

    /**
     * 是否隐藏
     */
    @JsonIgnore
    private boolean hidden;

    /**
     * 宽度，默认根据内容适应，Unit: a point
     */
    @JsonIgnore
    private Integer width;

    public ColumnDefinition() {

    }

    public ColumnDefinition(T sampleColumnModel) {
        this.sampleColumnModel = sampleColumnModel;
    }

    public ColumnDefinition(String name, String fieldNameOfValue, String fieldTypeOfValue) {
        this.name = name;
        this.fieldNameOfValue = fieldNameOfValue;
        this.fieldTypeOfValue = fieldTypeOfValue;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getFieldNameOfValue() {
        return fieldNameOfValue;
    }

    public void setFieldNameOfValue(String fieldNameOfValue) {
        this.fieldNameOfValue = fieldNameOfValue;
    }

    public String getFieldTypeOfValue() {
        if (fieldTypeOfValue == null
                && sampleColumnModel != null
                && !(sampleColumnModel instanceof Map)) {
            try {
                Field field = sampleColumnModel.getClass().getDeclaredField(fieldNameOfValue);
                fieldTypeOfValue = field.getClass().getName();
            } catch (Exception ex) {

            }
        }
        return fieldTypeOfValue;
    }

    public void setFieldTypeOfValue(String fieldTypeOfValue) {
        this.fieldTypeOfValue = fieldTypeOfValue;
    }

    public String getValueFormatPattern() {
        return valueFormatPattern;
    }

    public void setValueFormatPattern(String valueFormatPattern) {
        this.valueFormatPattern = valueFormatPattern;
    }

    public T getSampleColumnModel() {
        return sampleColumnModel;
    }

    public void setSampleColumnModel(T sampleColumnModel) {
        this.sampleColumnModel = sampleColumnModel;
    }

    public String getClassNameOfColumnModel() {
        if (sampleColumnModel != null) {
            classNameOfColumnModel = sampleColumnModel.getClass().getName();
        }

        return classNameOfColumnModel;
    }

    public Integer getDataColumnIndex() {
        return dataColumnIndex;
    }

    public void setDataColumnIndex(Integer dataColumnIndex) {
        this.dataColumnIndex = dataColumnIndex;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public Integer getFormulaColumnIndex() {
        return formulaColumnIndex;
    }

    public void setFormulaColumnIndex(Integer formulaColumnIndex) {
        this.formulaColumnIndex = formulaColumnIndex;
    }

    public boolean isEditable() {
        if (editable != null) {
            return editable;
        }

        return (getFormula() != null) ? false : true;
    }

    public void setEditable(boolean editable) {
        this.editable = editable;
    }

    public XSSFColor getReadOnlyCellColor() {
        return this.readOnlyCellColor;
    }

    public void setReadOnlyCellColor(XSSFColor readOnlyCellColor) {
        this.readOnlyCellColor = readOnlyCellColor;
    }

    public XSSFColor getEditableCellColor() {
        return this.editableCellColor;
    }

    public void setEditableCellColor(XSSFColor editableCellColor) {
        this.editableCellColor = editableCellColor;
    }

    public boolean isHidden() {
        return hidden;
    }

    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    @Override
    public String toString() {
        final StringBuilder sb = new StringBuilder("ColumnDefinition{");
        sb.append("fieldNameOfValue='").append(fieldNameOfValue).append('\'');
        sb.append(", valueFormatPattern='").append(valueFormatPattern).append('\'');
        sb.append(", sampleColumnModel=").append(sampleColumnModel);
        sb.append(", classNameOfColumnModel='").append(classNameOfColumnModel).append('\'');
        sb.append(", fieldTypeOfValue='").append(fieldTypeOfValue).append('\'');
        sb.append(", dataColumnIndex=").append(dataColumnIndex);
        sb.append(", formulaColumnIndex=").append(formulaColumnIndex);
        sb.append(", editable=").append(editable);
        sb.append(", readOnlyCellColor=").append(readOnlyCellColor);
        sb.append(", editableCellColor=").append(editableCellColor);
        sb.append(", hidden=").append(hidden);
        sb.append(", width=").append(width);
        sb.append(", ").append(super.toString());
        sb.append('}');
        return sb.toString();
    }
}
