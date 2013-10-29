package bc.bms.common.workbook.model;

import java.io.Serializable;

/**
 * 工作表配置模型
 */
public class WorksheetPreference implements Serializable {

    /**
     * 工作表配置占用行数
     */
    public static final int WORKSHEET_PREFERENCE_ROW_COUNT = 2;

    /**
     * 工作表配置存储起始列索引
     */
    public static final int WORKSHEET_PREFERENCE_START_COLUMN_INDEX = 300;

    /**
     * 工作表全局参数存储行索引
     */
    public static final int WORKSHEET_GLOBAL_PREFERENCE_ROW_INDEX = 0;

    /**
     * 工作表全局参数存储列索引
     */
    public static final int WORKSHEET_PREFERENCE_COLUMN_INDEX = WORKSHEET_PREFERENCE_START_COLUMN_INDEX;

    /**
     * 主键列定义数据存储行索引
     */
//    public static final int WORKSHEET_PRIMARY_COLUMN_PREFERENCE_ROW_INDEX = 1;

    /**
     * 列定义数据存储行索引
     */
    public static final int WORKSHEET_COLUMN_PREFERENCE_ROW_INDEX = 1;

    /**
     * 主键列索引
     */
    public static final int WORKSHEET_PRIMARY_COLUMN_INDEX = 0;

    /**
     * 是否支持导入
     */
    private boolean importable;

    /**
     * 表头占用行数
     */
    private int rowCountOfHeader;

    /**
     * 行高度，Unit: a point
     */
    private Short rowHeight;

    /**
     * 表头行高度，默认等于rowHeight，Unit: 1/20 of a point
     */
    private Short headerRowHeight;

    /**
     * 是否锁定
     */
    private boolean locked;

    public boolean isImportable() {
        return importable;
    }

    public void setImportable(boolean importable) {
        this.importable = importable;
    }

    public int getRowCountOfHeader() {
        return rowCountOfHeader;
    }

    public void setRowCountOfHeader(int rowCountOfHeader) {
        this.rowCountOfHeader = rowCountOfHeader;
    }

    public Short getRowHeight() {
        return rowHeight;
    }

    public void setRowHeight(Short rowHeight) {
        this.rowHeight = rowHeight;
    }

    public Short getHeaderRowHeight() {
        return headerRowHeight;
    }

    public void setHeaderRowHeight(Short headerRowHeight) {
        this.headerRowHeight = headerRowHeight;
    }

    public boolean isLocked() {
        return locked;
    }

    public void setLocked(boolean locked) {
        this.locked = locked;
    }

}
