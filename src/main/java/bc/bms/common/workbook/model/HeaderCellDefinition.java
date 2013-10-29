package bc.bms.common.workbook.model;

import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.Serializable;

/**
 * 表头单元格定义模型
 */
public class HeaderCellDefinition implements Serializable {

    /**
     * 表头内容
     */
    private String content;

    /**
     * 起始行
     */
    private int beginRow;

    /**
     * 结束行
     */
    private int endRow;

    /**
     * 起始列
     */
    private int beginColumn;

    /**
     * 结束列
     */
    private int endColumn;

    /**
     * 单元格背景色，参见{@link bc.bms.common.workbook.ColorPicker}
     */
    private XSSFColor cellColor;

    public HeaderCellDefinition() {

    }

    public HeaderCellDefinition(
            String content, int beginRow, int endRow, int beginColumn, int endColumn, XSSFColor cellColor) {
        this.content = content;
        this.beginRow = beginRow;
        this.endRow = endRow;
        this.beginColumn = beginColumn;
        this.endColumn = endColumn;
        this.cellColor = cellColor;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public int getBeginRow() {
        return beginRow;
    }

    public void setBeginRow(int beginRow) {
        this.beginRow = beginRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public int getBeginColumn() {
        return beginColumn;
    }

    public void setBeginColumn(int beginColumn) {
        this.beginColumn = beginColumn;
    }

    public int getEndColumn() {
        return endColumn;
    }

    public void setEndColumn(int endColumn) {
        this.endColumn = endColumn;
    }

    public XSSFColor getCellColor() {
        return cellColor;
    }

    public void setCellColor(XSSFColor cellColor) {
        this.cellColor = cellColor;
    }

}
