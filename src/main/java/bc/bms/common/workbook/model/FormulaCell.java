package bc.bms.common.workbook.model;

/**
 * 公式单元格模型
 */
public interface FormulaCell extends Editable {

    /**
     * 公式定义，仅支持四则运算
     * <pre>
     * 示例：
     * 1. (BC2 + BC3 - BC4) * BC5 / BC6
     * > BC代表公式是基于列之间的计算
     * > 数字代表列索引
     * 2. (BR2 + BR3 - BR4) * BR5 / BR6
     * > BR代表公式是基于行之间的计算
     * > 数字代表行索引
     * </pre>
     */
    public String getFormula();

    /**
     * 是否可编辑, {@code formula != null},  则不可编辑。
     *
     * @return
     */
    public boolean isEditable();

}
