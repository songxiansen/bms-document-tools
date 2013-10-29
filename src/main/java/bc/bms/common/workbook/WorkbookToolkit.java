package bc.bms.common.workbook;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;

/**
 * 常用工作簿处理工具
 */
public final class WorkbookToolkit {

    private WorkbookToolkit() {

    }

    /**
     * 获取或创建指定索引的行
     *
     * @param worksheet        目标工作表
     * @param rowIndex         行索引
     * @param createIfNotExist 行不存在时是否创建新的
     * @return 行对象
     */
    public static XSSFRow getRow(XSSFSheet worksheet, int rowIndex, boolean createIfNotExist) {
        XSSFRow row = worksheet.getRow(rowIndex);

        if (row == null && createIfNotExist) {
            row = worksheet.createRow(rowIndex);
        }

        return row;
    }

    /**
     * 获取指定行列的单元格，如果不存在则创建新单元格
     *
     * @param row              目标行
     * @param columnIndex      列索引
     * @param createIfNotExist 单元格不存在时是否创建新的
     * @return 单元格对象
     */
    public static XSSFCell getCell(XSSFRow row, int columnIndex, boolean createIfNotExist) {
        XSSFCell cell = row.getCell(columnIndex);

        if (cell == null && createIfNotExist) {
            cell = row.createCell(columnIndex);
        }

        return cell;
    }

    /**
     * 读取并返回工作簿
     *
     * @param path 工作簿路径
     * @return 工作簿对象
     */
    public static XSSFWorkbook readWorkbook(String path) {
        File file = new File(path);

        XSSFWorkbook workbook = readWorkbook(file);

        return workbook;
    }

    /**
     * 读取并返回工作簿
     *
     * @param file 工作簿File
     * @return 工作簿对象
     */
    public static XSSFWorkbook readWorkbook(File file) {
        XSSFWorkbook workbook = null;

        try {
            workbook = (XSSFWorkbook) WorkbookFactory.create(file);
        } catch (Exception ex) {
            throw new IllegalArgumentException("Excel 2007+ document is required.");
        }

        return workbook;
    }

    /**
     * 读取并返回工作簿
     *
     * @param inputStream 工作簿InputStream
     * @return 工作簿对象
     */
    public static XSSFWorkbook readWorkbook(InputStream inputStream) {
        XSSFWorkbook workbook = null;

        try {
            workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream);
        } catch (Exception ex) {
            throw new IllegalArgumentException("Excel 2007+ document is required.");
        }

        return workbook;
    }

}
