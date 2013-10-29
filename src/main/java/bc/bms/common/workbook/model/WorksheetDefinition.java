package bc.bms.common.workbook.model;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 工作表定义模型
 */
public class WorksheetDefinition implements Serializable {

    /**
     * 工作表名称
     */
    private String name;

    /**
     * 工作表表头定义
     */
    private List<HeaderCellDefinition> headerCellDefinitions;

    /**
     * 工作表列定义
     */
    private Map<Integer, ColumnDefinition> columnDefinitions;

    /**
     * 计算列单元格定义清单
     */
    private Map<Integer, List<CalculateColumnDefinition>> calculateColumnDefinitions;

    /**
     * 工作表设置
     */
    private WorksheetPreference worksheetPreference;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<HeaderCellDefinition> getHeaderCellDefinitions() {
        return headerCellDefinitions;
    }

    public void setHeaderCellDefinitions(List<HeaderCellDefinition> headerCellDefinitions) {
        this.headerCellDefinitions = headerCellDefinitions;
    }

    public Map<Integer, ColumnDefinition> getColumnDefinitions() {
        return columnDefinitions;
    }

    public void setColumnDefinitions(Map<Integer, ColumnDefinition> columnDefinitions) {
        this.columnDefinitions = columnDefinitions;
    }

    public void addColumnDefinition(Integer columnIndex, ColumnDefinition columnDefinition) {
        if (this.columnDefinitions == null) {
            this.columnDefinitions = new HashMap<Integer, ColumnDefinition>();
        }

        this.columnDefinitions.put(columnIndex, columnDefinition);
    }

    public Map<Integer, List<CalculateColumnDefinition>> getCalculateColumnDefinitions() {
        return calculateColumnDefinitions;
    }

    public void setCalculateColumnDefinitions(
            Map<Integer, List<CalculateColumnDefinition>> calculateColumnDefinitions) {
        this.calculateColumnDefinitions = calculateColumnDefinitions;
    }

    public static void addDataSetOfColumn(
            Map<Integer, List> dataSetOfColumns, Integer columnIndex, List dataSetOfColumn) {
        if (dataSetOfColumns == null) {
            dataSetOfColumns = new HashMap<Integer, List>();
        }

        dataSetOfColumns.put(columnIndex, dataSetOfColumn);
    }

    public static void addDataOfColumn(Map<Integer, List> dataSetOfColumns, Integer columnIndex, Object data) {
        if (dataSetOfColumns == null) {
            dataSetOfColumns = new HashMap<Integer, List>();
        }

        List<Object> dataSet = dataSetOfColumns.get(columnIndex);
        if (dataSetOfColumns.get(columnIndex) == null) {
            dataSet = new ArrayList<Object>();
            dataSetOfColumns.put(columnIndex, dataSet);
        }

        dataSet.add(data);
    }

    public WorksheetPreference getWorksheetPreference() {
        return worksheetPreference;
    }

    public void setWorksheetPreference(WorksheetPreference worksheetPreference) {
        this.worksheetPreference = worksheetPreference;
    }

}
