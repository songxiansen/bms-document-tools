package bc.bms.common.workbook.model;

/**
 * 用于确定数据是否可编辑，支持Excel导入导出功能的模型必须实现该接口
 */
public interface Editable {

    /**
     * 是否可编辑
     */
    public boolean isEditable();

}
