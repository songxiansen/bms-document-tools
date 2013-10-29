package bc.bms.common.workbook;

import bc.bms.common.util.JsonUtil;
import bc.bms.common.workbook.model.CalculateColumnDefinition;
import bc.bms.common.workbook.model.ColumnDefinition;
import bc.bms.common.workbook.model.Editable;
import bc.bms.common.workbook.model.FormulaCell;
import bc.bms.common.workbook.model.HeaderCellDefinition;
import bc.bms.common.workbook.model.WorksheetDefinition;
import bc.bms.common.workbook.model.WorksheetPreference;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 工作簿写入工具
 */
public final class WorkbookWriter {

    private static Log logger = LogFactory.getLog(WorkbookWriter.class);

    private WorkbookWriter() {

    }

    /**
     * 写入数据到工作簿
     *
     * @param workbook            目标工作簿，如果为null则创建新的工作簿
     * @param worksheetDefinition 要写入的工作表定义
     * @param dataSetOfColumns    要写入的数据集
     * @return 工作簿
     */
    public static XSSFWorkbook writeToWorkbook(
            XSSFWorkbook workbook, WorksheetDefinition worksheetDefinition, Map<Integer, List> dataSetOfColumns) {
        if (workbook == null) {
            workbook = new XSSFWorkbook();
        }

        XSSFSheet worksheet = workbook.createSheet(worksheetDefinition.getName());

        setDefaultStyle(worksheet, worksheetDefinition);

        populateWorksheetPreference(worksheet, worksheetDefinition);

        populateColumnDefinitions(worksheet, worksheetDefinition);

        populateWorksheetHeader(worksheet, worksheetDefinition);

        populateWorksheetData(worksheet, worksheetDefinition, dataSetOfColumns);

        if (worksheetDefinition.getWorksheetPreference().isLocked()) {
            worksheet.protectSheet("BMSExcelAdminPassword");
        }

        XSSFRow firstDataRow = worksheet.getRow(WorksheetPreference.WORKSHEET_PREFERENCE_ROW_COUNT);

        firstDataRow.getCell(0).setAsActiveCell();

        return workbook;
    }

    /**
     * 追加数据到指定工作表
     *
     * @param worksheet           目标工作表
     * @param worksheetDefinition 工作表定义数据
     * @param dataSetOfColumns    要写入的数据集
     */
    public static void appendDataToWorksheet(
            XSSFSheet worksheet, WorksheetDefinition worksheetDefinition, Map<Integer, List> dataSetOfColumns) {
        populateWorksheetData(worksheet, worksheetDefinition, dataSetOfColumns);
    }

    /**
     * 设置默认样式
     *
     * @param worksheet           目标工作表
     * @param worksheetDefinition 工作表定义数据
     */
    private static void setDefaultStyle(XSSFSheet worksheet, WorksheetDefinition worksheetDefinition) {
        XSSFCellStyle defaultColumnStyle = worksheet.getWorkbook().createCellStyle();
        defaultColumnStyle.setAlignment(CellStyle.ALIGN_LEFT);
        defaultColumnStyle.setVerticalAlignment(CellStyle.VERTICAL_TOP);
        defaultColumnStyle.setBorderBottom(CellStyle.BORDER_THIN);
        defaultColumnStyle.setBorderTop(CellStyle.BORDER_THIN);
        defaultColumnStyle.setBorderRight(CellStyle.BORDER_THIN);
        defaultColumnStyle.setBorderLeft(CellStyle.BORDER_THIN);

        Font font = Font.defaultFont(worksheet.getWorkbook());
        defaultColumnStyle.setFont(font);

        for (Map.Entry<Integer, ColumnDefinition> columnDefinitionEntry
                : worksheetDefinition.getColumnDefinitions().entrySet()) {
            int columnIndex = columnDefinitionEntry.getKey();
            ColumnDefinition columnDefinition = columnDefinitionEntry.getValue();

            CellStyle currentColumnStyle = worksheet.getWorkbook().createCellStyle();
            currentColumnStyle.cloneStyleFrom(defaultColumnStyle);
            worksheet.setDefaultColumnStyle(columnIndex, currentColumnStyle);

            if (columnDefinition.getWidth() != null) {
                worksheet.setColumnWidth(columnIndex, columnDefinition.getWidth() * 20);
            } else {
                worksheet.setColumnWidth(columnIndex, 4000);
            }

            if (columnDefinition.isHidden()) {
                worksheet.setColumnHidden(columnIndex, true);
            }
        }
    }

    /**
     * 写入工作表配置数据
     *
     * @param worksheet           目标工作表
     * @param worksheetDefinition 工作表定义数据
     */
    private static void populateWorksheetPreference(
            XSSFSheet worksheet, WorksheetDefinition worksheetDefinition) {
        XSSFRow worksheetPreferenceRow
                = WorkbookToolkit.getRow(worksheet, WorksheetPreference.WORKSHEET_GLOBAL_PREFERENCE_ROW_INDEX, true);

        worksheetPreferenceRow.setZeroHeight(true);

        XSSFCell worksheetPreferenceCell = WorkbookToolkit.getCell(
                worksheetPreferenceRow,
                WorksheetPreference.WORKSHEET_PREFERENCE_COLUMN_INDEX, true);

        int rowCountOfHeader = 1;
        for (HeaderCellDefinition headerCellDefinition : worksheetDefinition.getHeaderCellDefinitions()) {
            if (headerCellDefinition.getEndRow() + 1 > rowCountOfHeader) {
                rowCountOfHeader = headerCellDefinition.getEndRow() + 1;
            }
        }

        worksheetDefinition.getWorksheetPreference().setRowCountOfHeader(rowCountOfHeader);

        worksheetPreferenceCell.setCellValue(JsonUtil.toJson(worksheetDefinition.getWorksheetPreference()));
    }

    /**
     * 写入列定义数据到对应工作表
     *
     * @param worksheet           要写入的工作表
     * @param worksheetDefinition 工作表定义及数据
     */
    private static void populateColumnDefinitions(
            XSSFSheet worksheet, WorksheetDefinition worksheetDefinition) {
        XSSFRow rowOfColumnDefinitions
                = WorkbookToolkit.getRow(worksheet, WorksheetPreference.WORKSHEET_COLUMN_PREFERENCE_ROW_INDEX, true);

        XSSFCellStyle columnDefinitionCellStyle = worksheet.getWorkbook().createCellStyle();
        columnDefinitionCellStyle.setWrapText(true);
        columnDefinitionCellStyle.setHidden(true);

        for (Map.Entry<Integer, ColumnDefinition> columnDefinitionEntry
                : worksheetDefinition.getColumnDefinitions().entrySet()) {
            int columnIndex = columnDefinitionEntry.getKey();
            ColumnDefinition columnDefinition = columnDefinitionEntry.getValue();

            XSSFCell cell = rowOfColumnDefinitions.createCell(columnIndex
                    + WorksheetPreference.WORKSHEET_PREFERENCE_START_COLUMN_INDEX);
            cell.setCellValue(JsonUtil.toJson(columnDefinition).toString());

            cell.setCellStyle(columnDefinitionCellStyle);
        }

        rowOfColumnDefinitions.setZeroHeight(true);
    }

    /**
     * 写入工作表头
     *
     * @param worksheet           要写入的工作表
     * @param worksheetDefinition 工作表定义及数据
     */
    private static void populateWorksheetHeader(XSSFSheet worksheet, WorksheetDefinition worksheetDefinition) {
        int rowCountOfHeader = worksheetDefinition.getWorksheetPreference().getRowCountOfHeader();

        for (int i = 0; i < rowCountOfHeader; i++) {
            WorkbookToolkit.getRow(worksheet, WorksheetPreference.WORKSHEET_PREFERENCE_ROW_COUNT + i, true);
        }

        XSSFCellStyle cellStyle = worksheet.getWorkbook().createCellStyle();

        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(ColorPicker.getColor("#64B664"));

        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);

        Font font = Font.defaultHeaderFont(worksheet.getWorkbook());
        cellStyle.setFont(font);

        for (HeaderCellDefinition headerCellDefinition : worksheetDefinition.getHeaderCellDefinitions()) {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(
                    headerCellDefinition.getBeginRow() + WorksheetPreference.WORKSHEET_PREFERENCE_ROW_COUNT,
                    headerCellDefinition.getEndRow() + WorksheetPreference.WORKSHEET_PREFERENCE_ROW_COUNT,
                    headerCellDefinition.getBeginColumn(), headerCellDefinition.getEndColumn());
            worksheet.addMergedRegion(cellRangeAddress);

            for (int i = headerCellDefinition.getBeginRow(); i <= headerCellDefinition.getEndRow(); i++) {
                XSSFRow headerRow = WorkbookToolkit.getRow(
                        worksheet, i + WorksheetPreference.WORKSHEET_PREFERENCE_ROW_COUNT,
                        true);

                if (worksheetDefinition.getWorksheetPreference().getHeaderRowHeight() != null) {
                    headerRow.setHeightInPoints(worksheetDefinition.getWorksheetPreference().getHeaderRowHeight());
                }

                for (int j = headerCellDefinition.getBeginColumn(); j <= headerCellDefinition.getEndColumn(); j++) {
                    XSSFCell headerRowCell = WorkbookToolkit.getCell(headerRow, j, true);
                    headerRowCell.setCellType(XSSFCell.CELL_TYPE_STRING);
                    headerRowCell.setCellValue(headerCellDefinition.getContent());

                    CellStyle currentCellStyle = worksheet.getWorkbook().createCellStyle();
                    currentCellStyle.cloneStyleFrom(cellStyle);
                    headerRowCell.setCellStyle(currentCellStyle);

                    if (headerCellDefinition.getCellColor() != null) {
                        headerRowCell.getCellStyle().setFillForegroundColor(headerCellDefinition.getCellColor());
                    }
                }
            }
        }
    }

    /**
     * 填充数据区域
     *
     * @param worksheet           要写入的工作表
     * @param worksheetDefinition 工作表定义及数据
     * @param dataSetOfColumns    要写入的数据集
     */
    private static void populateWorksheetData(
            XSSFSheet worksheet, WorksheetDefinition worksheetDefinition, Map<Integer, List> dataSetOfColumns) {
        int startRowIndex = worksheet.getLastRowNum() + 1;

        if (dataSetOfColumns == null || dataSetOfColumns.size() == 0) {
            logger.info("No data need to export.");
            return;
        }

        populateWorksheetData(worksheet, startRowIndex, worksheetDefinition, dataSetOfColumns);
    }

    /**
     * 填充数据区域
     *
     * @param worksheet           要写入的工作表
     * @param startRowIndex       工作表数据区域开始行索引
     * @param worksheetDefinition 工作表定义及数据
     * @param dataSetOfColumns    要写入的数据集
     */
    private static void populateWorksheetData(
            XSSFSheet worksheet, int startRowIndex, WorksheetDefinition worksheetDefinition,
            Map<Integer, List> dataSetOfColumns) {
        try {
            int rowCount = dataSetOfColumns.values().iterator().next().size();
            for (int i = 0; i < rowCount; i++) {
                int rowIndex = startRowIndex + i;
                XSSFRow row = WorkbookToolkit.getRow(worksheet, rowIndex, true);

                for (Map.Entry<Integer, ColumnDefinition>
                        columnDefinitionEntry : worksheetDefinition.getColumnDefinitions().entrySet()) {
                    Integer columnIndex = columnDefinitionEntry.getKey();
                    ColumnDefinition columnDefinition = columnDefinitionEntry.getValue();

                    List<Object> dataSetOfColumn = null;

                    try {
                        dataSetOfColumn = dataSetOfColumns.get(columnIndex);
                    } catch (Exception ex) {
                    }

                    if (dataSetOfColumn == null && columnDefinition.getDataColumnIndex() != null) {
                        dataSetOfColumn = dataSetOfColumns.get(columnDefinition.getDataColumnIndex());
                    }

                    if (dataSetOfColumn == null) {
                        logger.warn("Can't get dataSet of column[" + columnIndex + "].");
                        continue;
                    }

                    if (columnIndex == WorksheetPreference.WORKSHEET_PRIMARY_COLUMN_INDEX) {
                        if (worksheetDefinition.getWorksheetPreference().getRowHeight() != null) {
                            row.setHeightInPoints(worksheetDefinition.getWorksheetPreference().getRowHeight());
                        }
                    }

                    XSSFCell cell = WorkbookToolkit.getCell(row, columnIndex, true);

                    populateWorksheetDataByType(
                            worksheetDefinition, dataSetOfColumn.get(i), columnDefinition, cell);
                }

                logger.info("Populate data into row[" + rowIndex + "] finished.");
            }
        } catch (Exception exception) {
            throw new IllegalArgumentException(exception.getMessage(), exception);
        }
    }

    /**
     * 根据类型写入数据到指定单元格
     *
     * @param worksheetDefinition 工作表定义及数据
     * @param data                要写入的数据
     * @param columnDefinition    列定义数据
     * @param cell                目标单元格
     * @throws IllegalArgumentException 发生错误时抛出无效参数异常
     */
    private static void populateWorksheetDataByType(
            WorksheetDefinition worksheetDefinition, Object data,
            ColumnDefinition columnDefinition, XSSFCell cell)
            throws IllegalArgumentException {
        try {
            if (data instanceof Map) {
                populateDataInMap(worksheetDefinition, data, columnDefinition, cell);
            } else {
                populateDataOfCustomizedType(worksheetDefinition, data, columnDefinition, cell);
            }

            if (columnDefinition.getReadOnlyCellColor() != null) {
                cell.getCellStyle().setFillForegroundColor(columnDefinition.getReadOnlyCellColor());
            }

            boolean editable = false;
            if (data instanceof Editable) {
                Editable editableData = (Editable) data;
                editable = editableData.isEditable();
            } else if (columnDefinition.isEditable()) {
                editable = true;
            }

            if (editable) {
                cell.getCellStyle().setLocked(false);

                if (columnDefinition.getEditableCellColor() != null) {
                    cell.getCellStyle().setFillForegroundColor(columnDefinition.getEditableCellColor());
                }
            }
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 写入Map格式数据
     *
     * @param worksheetDefinition 工作表定义及数据
     * @param data                要写入的数据
     * @param columnDefinition    列定义数据
     * @param cell                目标单元格
     * @throws IllegalArgumentException 发生错误时抛出无效参数异常
     */
    private static void populateDataInMap(
            WorksheetDefinition worksheetDefinition, Object data, ColumnDefinition columnDefinition, XSSFCell cell)
            throws IllegalArgumentException {
        try {
            Map<String, Object> dataInMap = (Map<String, Object>) data;
            Object valueObject = dataInMap.get(columnDefinition.getFieldNameOfValue());

            List<CalculateColumnDefinition> calculateColumnDefinitions = null;

            if (worksheetDefinition.getCalculateColumnDefinitions() != null) {
                calculateColumnDefinitions = worksheetDefinition
                        .getCalculateColumnDefinitions().get(columnDefinition.getFormulaColumnIndex());
            }

            populateData(valueObject, columnDefinition, cell, calculateColumnDefinitions);
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 写入自定义格式数据
     *
     * @param worksheetDefinition 工作表定义及数据
     * @param data                要写入的数据
     * @param columnDefinition    列定义数据
     * @param cell                目标单元格
     * @throws IllegalAccessException 发生错误时抛出无效参数异常
     */
    private static void populateDataOfCustomizedType(
            WorksheetDefinition worksheetDefinition, Object data, ColumnDefinition columnDefinition, XSSFCell cell)
            throws IllegalAccessException {
        try {
            Object valueObject = PropertyUtils.getProperty(data, columnDefinition.getFieldNameOfValue());

            List<CalculateColumnDefinition> calculateColumnDefinitions = null;

            if (worksheetDefinition.getCalculateColumnDefinitions() != null) {
                calculateColumnDefinitions = worksheetDefinition
                        .getCalculateColumnDefinitions().get(columnDefinition.getFormulaColumnIndex());
            }

            populateData(valueObject, columnDefinition, cell, calculateColumnDefinitions);
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 写入单元格数据
     *
     * @param valueObject                要写入的单元格数据
     * @param cell                       目标单元格
     * @param calculateColumnDefinitions 计算列单元格定义清单
     */
    private static void populateData(
            Object valueObject, ColumnDefinition columnDefinition, XSSFCell cell,
            List<CalculateColumnDefinition> calculateColumnDefinitions) {
        String originalFormula = null;

        if (columnDefinition.getFormula() != null) {
            originalFormula = columnDefinition.getFormula();
        }

        if (columnDefinition.getFormulaColumnIndex() != null) {
            if (calculateColumnDefinitions == null || calculateColumnDefinitions.get(cell.getRowIndex()) == null) {
                logger.warn("The definitions of calculate column is required," +
                        " when using the formulation from other column.");
            } else if (calculateColumnDefinitions.get(cell.getRowIndex()).getFormula() != null) {
                originalFormula = calculateColumnDefinitions.get(cell.getRowIndex()).getFormula();
            }
        }

        if (valueObject != null && valueObject instanceof FormulaCell) {
            FormulaCell formulaCell = (FormulaCell) valueObject;

            if (formulaCell.getFormula() != null) {
                originalFormula = formulaCell.getFormula();
            }
        }

        if (originalFormula != null) {
            cell.setCellType(Cell.CELL_TYPE_FORMULA);
            cell.setCellFormula(getFormula(originalFormula, cell));
        } else if (valueObject != null) {
            if (valueObject instanceof BigDecimal) {
                BigDecimal value = (BigDecimal) valueObject;
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue(value.doubleValue());
            } else if (valueObject instanceof Integer) {
                Integer value = (Integer) valueObject;
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue(value.doubleValue());
            } else if (valueObject instanceof String) {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue(valueObject.toString());
            } else {
                throw new IllegalArgumentException("Unsupported Type "
                        + valueObject.getClass());
            }

            if (valueObject instanceof Number) {
                cell.getCellStyle().setAlignment(CellStyle.ALIGN_RIGHT);
            }
        }
    }

    /**
     * 解析公式
     *
     * @param originalFormula 原始公式
     * @param cell          目标单元格
     * @return 翻译之后的公式
     */
    private static String getFormula(String originalFormula, XSSFCell cell) {
        String formula = originalFormula;

        if (formula.contains("BC")) {
            try {
                List<String> matchedStrings = new ArrayList<String>();

                Pattern pattern = Pattern.compile("BC[0-9]+");
                Matcher matcher = pattern.matcher(formula);
                while (matcher.find()) {
                    matchedStrings.add(matcher.group());
                }

                for (String matchedString : matchedStrings) {
                    XSSFRow row = cell.getRow();
                    XSSFCell referenceCell = WorkbookToolkit.getCell(
                            row, Integer.parseInt(matchedString.replace("BC", "")), true);

                    formula = formula.replaceAll(matchedString, referenceCell.getReference());
                }
            } catch (Exception ex) {
                throw new IllegalArgumentException(ex);
            }
        } else if (formula.contains("BR")) {
            formula = formula.replaceAll("BR", cell.getReference());
        }

        return formula;
    }

}
