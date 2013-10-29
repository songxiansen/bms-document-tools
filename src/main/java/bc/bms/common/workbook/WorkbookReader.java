package bc.bms.common.workbook;

import bc.bms.common.util.JsonUtil;
import bc.bms.common.workbook.model.CalculateColumnDefinition;
import bc.bms.common.workbook.model.ColumnDefinition;
import bc.bms.common.workbook.model.WorksheetPreference;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.JsonNode;

import java.io.File;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * 工作簿读取工具
 */
public final class WorkbookReader {

    private static Log logger = LogFactory.getLog(WorkbookReader.class);

    private WorkbookReader() {

    }

    /**
     * 通过模板读取工作簿定义数据
     *
     * @param path                          模板文件路径
     * @param sampleCalculateColumnDefinition CalculateColumnDefinition示例
     * @param fieldNames                    存储值的Field names, 顺序与工作表列索引保持一致
     * @return 工作簿名称及其对应的定义数据
     */
    public static <T> Map<String, List<CalculateColumnDefinition<T>>> getCalculateColumnDefinitionFromTemplate(
            String path, CalculateColumnDefinition<T> sampleCalculateColumnDefinition, String[] fieldNames) {
        XSSFWorkbook workbook = WorkbookToolkit.readWorkbook(path);

        return getCalculateColumnDefinitionFromTemplate(workbook, sampleCalculateColumnDefinition, fieldNames);
    }

    /**
     * 通过模板读取工作簿定义数据
     *
     * @param file                          模板文件
     * @param sampleCalculateColumnDefinition CalculateColumnDefinition示例
     * @param fieldNames                    存储值的Field names, 顺序与工作表列索引保持一致
     * @return 工作簿名称及其对应的定义数据
     */
    public static <T> Map<String, List<CalculateColumnDefinition<T>>> getCalculateColumnDefinitionFromTemplate(
            File file, CalculateColumnDefinition<T> sampleCalculateColumnDefinition, String[] fieldNames) {
        XSSFWorkbook workbook = WorkbookToolkit.readWorkbook(file);

        return getCalculateColumnDefinitionFromTemplate(workbook, sampleCalculateColumnDefinition, fieldNames);
    }

    /**
     * 通过模板读取工作簿定义数据
     *
     * @param inputStream                   模板文件InputStream
     * @param sampleCalculateColumnDefinition CalculateColumnDefinition示例
     * @param fieldNames                    存储值的Field names, 顺序与工作表列索引保持一致
     * @return 工作簿名称及其对应的定义数据
     */
    public static <T> Map<String, List<CalculateColumnDefinition<T>>> getCalculateColumnDefinitionFromTemplate(
            InputStream inputStream, CalculateColumnDefinition<T> sampleCalculateColumnDefinition,
            String[] fieldNames) {
        XSSFWorkbook workbook = WorkbookToolkit.readWorkbook(inputStream);

        return getCalculateColumnDefinitionFromTemplate(workbook, sampleCalculateColumnDefinition, fieldNames);
    }

    /**
     * 读取工作簿定义数据
     *
     * @param workbook                      源工作簿
     * @param sampleCalculateColumnDefinition CalculateColumnDefinition示例
     * @param fieldNames                    存储值的Field names, 顺序与工作表列索引保持一致
     * @param <T>                           数据模型
     * @return 工作簿名称及其对应的定义数据
     */
    private static <T> Map<String, List<CalculateColumnDefinition<T>>> getCalculateColumnDefinitionFromTemplate(
            XSSFWorkbook workbook, CalculateColumnDefinition<T> sampleCalculateColumnDefinition,
            String[] fieldNames) {
        Map<String, List<CalculateColumnDefinition<T>>> calculateColumnDefinitionsSeparatedByWorksheet
                = new HashMap<String, List<CalculateColumnDefinition<T>>>();

        for (XSSFSheet worksheet : workbook) {
            List<CalculateColumnDefinition<T>>
                    calculateDefinitionsSeparatedByColumnIndex = getCalculateColumnDefinitionFromWorksheet(
                    worksheet, sampleCalculateColumnDefinition, fieldNames);

            if (calculateDefinitionsSeparatedByColumnIndex != null
                    && calculateDefinitionsSeparatedByColumnIndex.size() > 0) {
                calculateColumnDefinitionsSeparatedByWorksheet.put(
                        worksheet.getSheetName(), calculateDefinitionsSeparatedByColumnIndex);
            }
        }

        return calculateColumnDefinitionsSeparatedByWorksheet;
    }

    /**
     * 读取工作表定义数据
     *
     * @param worksheet                     源工作表
     * @param sampleCalculateColumnDefinition 数据模型示例
     * @param fieldNames                    存储值的Field names, 顺序与工作表列索引保持一致
     * @param <T>                           数据模型
     * @return 按列分组的工作表定义数据
     */
    private static <T> List<CalculateColumnDefinition<T>> getCalculateColumnDefinitionFromWorksheet(
            XSSFSheet worksheet, CalculateColumnDefinition<T> sampleCalculateColumnDefinition, String[] fieldNames) {
        List<CalculateColumnDefinition<T>> calculateDefinitionsSeparatedByColumnIndex
                = new ArrayList<CalculateColumnDefinition<T>>();

        Iterator<Row> rowIterator = worksheet.rowIterator();

        //Skip the first row
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        outer:
        while (rowIterator.hasNext()) {
            try {
                XSSFRow row = (XSSFRow) rowIterator.next();
                CalculateColumnDefinition<T> calculateColumnDefinition
                        = copyCalculateColumnDefinition(sampleCalculateColumnDefinition);

                for (int i = 0; i <= fieldNames.length; i++) {
                    XSSFCell cell = WorkbookToolkit.getCell(row, i, false);
                    String cellValue = cell.getRawValue();

                    if (i == 0 && (cellValue == null || cellValue.equals(""))) {
                        break outer;
                    }

                    if (i == fieldNames.length && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                        if (cell.getCellFormula() != null) {
                            calculateColumnDefinition.setFormula(cell.getCellFormula().replaceAll("[A-Z]+", "BR"));
                        }
                    } else if (i < fieldNames.length) {
                        setFieldValue(calculateColumnDefinition, fieldNames[i], cell);
                    }
                }

                calculateDefinitionsSeparatedByColumnIndex.add(calculateColumnDefinition);
            } catch (Exception ex) {
                throw new IllegalArgumentException(ex);
            }
        }

        return calculateDefinitionsSeparatedByColumnIndex;
    }

    /**
     * 复制主键列定义示例到新的对象
     *
     * @param sampleCalculateColumnDefinition 主键列定义示例
     * @return 新的主键列定义
     */
    private static CalculateColumnDefinition copyCalculateColumnDefinition(
            CalculateColumnDefinition sampleCalculateColumnDefinition) {
        try {
            CalculateColumnDefinition calculateColumnDefinition
                    = (CalculateColumnDefinition) BeanUtils.cloneBean(sampleCalculateColumnDefinition);
            calculateColumnDefinition
                    .setSampleColumnModel((BeanUtils.cloneBean(sampleCalculateColumnDefinition.getSampleColumnModel())));

            return calculateColumnDefinition;
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 复制单元格的值到主键列对象对应的Field
     *
     * @param calculateColumnDefinition 主键列定义
     * @param fieldName               目标Field
     * @param cell                    源单元格
     */
    private static void setFieldValue(
            CalculateColumnDefinition calculateColumnDefinition, String fieldName, XSSFCell cell) {
        try {
            Object sampleDataModel = calculateColumnDefinition.getSampleColumnModel();
            Field field = sampleDataModel.getClass().getDeclaredField(fieldName);

            setFieldValue(sampleDataModel, field, calculateColumnDefinition, cell);
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 从Excel工作簿提取数据
     *
     * @param path 源工作簿路径
     * @return 按列分组的数据集
     */
    public static Map<Integer, List<Object>> collectDataFromWorkbook(String path) {
        XSSFWorkbook workbook = WorkbookToolkit.readWorkbook(path);

        return collectData(workbook);
    }

    /**
     * 从Excel工作簿提取数据
     *
     * @param file 源工作簿File
     * @return 按列分组的数据集
     */
    public static Map<Integer, List<Object>> collectDataFromWorkbook(File file) {
        XSSFWorkbook workbook = WorkbookToolkit.readWorkbook(file);

        return collectData(workbook);
    }

    /**
     * 从Excel工作簿提取数据
     *
     * @param inputStream 源工作簿InputStream
     * @return 按列分组的数据集
     */
    public static Map<Integer, List<Object>> collectDataFromWorkbook(
            InputStream inputStream) {
        XSSFWorkbook workbook = WorkbookToolkit.readWorkbook(inputStream);

        return collectData(workbook);
    }

    /**
     * 从Excel工作簿提取数据
     *
     * @param workbook 源工作簿
     * @return 按列分组的数据集
     */
    private static Map<Integer, List<Object>> collectData(XSSFWorkbook workbook) {
        Map<Integer, List<Object>> dataSetSeparatedByColumnIndex = new HashMap<Integer, List<Object>>();

        Iterator<XSSFSheet> worksheets = workbook.iterator();

        while (worksheets.hasNext()) {
            XSSFSheet worksheet = worksheets.next();

            WorksheetPreference worksheetPreference = collectWorksheetPreference(worksheet);

            if (worksheetPreference != null && worksheetPreference.isImportable()) {
                collectData(worksheet, dataSetSeparatedByColumnIndex, worksheetPreference);
            } else {
                logger.info("Worksheet[" + worksheet.getSheetName() + "] is not importable, skipped.");
            }
        }

        return dataSetSeparatedByColumnIndex;
    }

    /**
     * 提取工作表配置
     *
     * @param worksheet 源工作表
     * @return 工作表配置
     */
    private static WorksheetPreference collectWorksheetPreference(XSSFSheet worksheet) {
        WorksheetPreference worksheetPreference = null;

        try {
            XSSFRow worksheetPreferenceRow = WorkbookToolkit.getRow(
                    worksheet, WorksheetPreference.WORKSHEET_GLOBAL_PREFERENCE_ROW_INDEX, false);

            XSSFCell worksheetPreferenceCell = WorkbookToolkit.getCell(
                    worksheetPreferenceRow,
                    WorksheetPreference.WORKSHEET_PREFERENCE_COLUMN_INDEX, false);

            worksheetPreference
                    = JsonUtil.toBean(worksheetPreferenceCell.getStringCellValue(), WorksheetPreference.class);
        } catch (Exception ex) {
            logger.warn("Can't get preference of worksheet");
        }

        return worksheetPreference;
    }

    /**
     * 从工作簿提取数据
     *
     * @param worksheet                     源工作表
     * @param dataSetSeparatedByColumnIndex 按列分组的数据集
     * @param worksheetPreference           工作表配置
     */
    private static void collectData(
            XSSFSheet worksheet, Map<Integer, List<Object>> dataSetSeparatedByColumnIndex,
            WorksheetPreference worksheetPreference) {
        try {
            Map<Integer, ColumnDefinition> columnDefinitions = collectColumnDefinitions(worksheet);

            for (Map.Entry<Integer, ColumnDefinition> columnDefinitionEntry : columnDefinitions.entrySet()) {
                int columnIndex = columnDefinitionEntry.getKey();
                ColumnDefinition columnDefinition = columnDefinitionEntry.getValue();

                if (columnDefinition.isEditable()) {
                    int rowIndex = WorksheetPreference.WORKSHEET_PREFERENCE_ROW_COUNT
                            + worksheetPreference.getRowCountOfHeader();
                    int lastRowIndex = worksheet.getLastRowNum();
                    while(rowIndex <= lastRowIndex) {
                        collectData(columnIndex, rowIndex, worksheet, worksheetPreference,
                                columnDefinition, dataSetSeparatedByColumnIndex);

                        rowIndex++;
                    }
                }
            }
        } catch (Exception ex) {
            logger.warn("Invalid worksheet", ex);
        }
    }

    /**
     * 提取列定义数据
     *
     * @param worksheet 源工作表
     * @return 列定义数据
     */
    private static Map<Integer, ColumnDefinition> collectColumnDefinitions(XSSFSheet worksheet) {
        Map<Integer, ColumnDefinition> columnDefinitions = new HashMap<Integer, ColumnDefinition>();

        try {
            XSSFRow columnDefinitionRow = WorkbookToolkit.getRow(
                    worksheet, WorksheetPreference.WORKSHEET_COLUMN_PREFERENCE_ROW_INDEX, false);

            int columnIndex = 0;
            while (true) {
                XSSFCell columnDefinitionCell = WorkbookToolkit.getCell(
                        columnDefinitionRow,
                        columnIndex + WorksheetPreference.WORKSHEET_PREFERENCE_START_COLUMN_INDEX, false);

                if (columnDefinitionCell == null) {
                    break;
                } else {
                    String cellValue = columnDefinitionCell.getStringCellValue();
                    ColumnDefinition columnDefinition = JsonUtil.toBean(
                            columnDefinitionCell.getStringCellValue(),
                            ColumnDefinition.class,
                            getClassOfColumnModel(cellValue));

                    if (columnDefinition != null) {
                        columnDefinitions.put(columnIndex, columnDefinition);
                    }
                }

                columnIndex++;
            }
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }

        return columnDefinitions;
    }

    /**
     * 从工作簿提取数据
     *
     * @param columnDefinitionInJsonFormat Json格式的列定义数据
     * @return 列模型
     * @throws ClassNotFoundException 模型无法确定
     */
    private static Class getClassOfColumnModel(String columnDefinitionInJsonFormat) throws ClassNotFoundException {
        JsonNode cellValueInJson = JsonUtil.toJsonNode(columnDefinitionInJsonFormat);
        JsonNode classNameOfColumnModelNode
                = cellValueInJson.get(ColumnDefinition.CLASS_NAME_OF_COLUMN_MODEL_FIELD_NAME);
        String classNameOfColumnModel = classNameOfColumnModelNode.getTextValue();
        Class classOfColumnModel = Class.forName(classNameOfColumnModel);

        return classOfColumnModel;
    }

    /**
     * 从工作簿提取数据
     *
     * @param columnIndex                   列索引
     * @param rowIndex                      行索引
     * @param worksheet                     源工作表
     * @param columnDefinition              列定义
     * @param dataSetSeparatedByColumnIndex 按类型分组的数据集
     */
    private static void collectData(
            int columnIndex, int rowIndex, XSSFSheet worksheet,
            WorksheetPreference worksheetPreference,
            ColumnDefinition columnDefinition, Map<Integer, List<Object>> dataSetSeparatedByColumnIndex) {
        XSSFRow dataRow = WorkbookToolkit.getRow(worksheet, rowIndex, false);
        XSSFCell dataCell = WorkbookToolkit.getCell(dataRow, columnIndex, false);

        try {
            Class columnModel = Class.forName(columnDefinition.getClassNameOfColumnModel());

            Object data;
            int dataColumnIndex = columnIndex;
            int dataRowIndex = rowIndex - WorksheetPreference.WORKSHEET_PREFERENCE_ROW_COUNT
                    - worksheetPreference.getRowCountOfHeader();
            if (columnDefinition.getDataColumnIndex() != null) {
                dataColumnIndex = columnDefinition.getDataColumnIndex();
            }

            List<Object> dataSet = dataSetSeparatedByColumnIndex.get(dataColumnIndex);
            if (dataSet == null) {
                dataSet = new ArrayList<Object>();
                dataSetSeparatedByColumnIndex.put(dataColumnIndex, dataSet);
            }

            if (dataSet.size() <= dataRowIndex) {
                data = BeanUtils.cloneBean(columnDefinition.getSampleColumnModel());
                dataSet.add(data);
            } else {
                data = dataSet.get(dataRowIndex);
            }

            if (data instanceof Map) {
                collectDataInMap(data, columnDefinition, dataCell);
            } else {
                collectDataOfCustomizedType(data, columnModel, columnDefinition, dataCell);
            }
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 提取以Map格式数据
     *
     * @param data                    单元格数据
     * @param columnDefinition        列定义
     * @param dataCell                单元格
     * @throws IllegalArgumentException 发生错误时返回无效参数异常
     */
    private static void collectDataInMap(Object data, ColumnDefinition columnDefinition, XSSFCell dataCell)
            throws IllegalAccessException, NoSuchMethodException, InvocationTargetException,
            NoSuchFieldException, ClassNotFoundException {
        try {
            Map<String, Object> dataInMap = (Map<String, Object>) data;

            Class valueType = Class.forName(columnDefinition.getFieldTypeOfValue());
            Object value = null;

            if (valueType.isAssignableFrom(BigDecimal.class)) {
                DecimalFormat decimalFormat = new DecimalFormat("0.00");
                if (columnDefinition.getValueFormatPattern() != null) {
                    decimalFormat = new DecimalFormat(columnDefinition.getValueFormatPattern());
                }

                try {
                    value = new BigDecimal(decimalFormat.format(dataCell.getNumericCellValue()));
                } catch (Exception ex) {
                    try {
                        value = new BigDecimal(dataCell.getStringCellValue());
                    } catch (Exception e) {
                    }
                }
            } else if (valueType.isAssignableFrom(Integer.class)) {
                try {
                    Double originalValue = Double.valueOf(dataCell.getNumericCellValue());
                    value = originalValue.intValue();
                } catch (Exception ex) {
                    try {
                        String originalValue = String.valueOf(dataCell.getNumericCellValue());
                        if ("是".equals(originalValue)) {
                            value = 1;
                        } else if ("否".equals(originalValue)) {
                            value = 0;
                        }
                    } catch (Exception e) {
                    }
                }
            } else if (valueType.isAssignableFrom(String.class)) {
                try {
                    value = dataCell.getStringCellValue();
                } catch (Exception ex) {
                    try {
                        value = String.valueOf(dataCell.getNumericCellValue());
                    } catch (Exception e) {
                    }
                }
            }

            dataInMap.put(columnDefinition.getFieldNameOfValue(), value);
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 提取自定义类型数据
     *
     * @param data                    单元格数据
     * @param columnModel             列模型
     * @param columnDefinition        列定义
     * @param dataCell                单元格
     * @throws IllegalArgumentException 发生错误时返回无效参数异常
     */
    private static void collectDataOfCustomizedType(
            Object data, Class columnModel, ColumnDefinition columnDefinition, XSSFCell dataCell)
            throws IllegalArgumentException {
        try {
            Field valueField = columnModel.getDeclaredField(columnDefinition.getFieldNameOfValue());

            setFieldValue(data, valueField, columnDefinition, dataCell);
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 复制单元格数据到指定数据对象的Field
     *
     * @param data             数据对象
     * @param valueField       目标Feild
     * @param columnDefinition 数据对象对应的列定义
     * @param dataCell         源单元格
     * @throws IllegalArgumentException 发生错误时返回无效参数异常
     */
    private static void setFieldValue(
            Object data, Field valueField, ColumnDefinition columnDefinition, XSSFCell dataCell)
            throws IllegalArgumentException {
        try {
            Object value = null;

            if (valueField.getType().isAssignableFrom(BigDecimal.class)) {
                DecimalFormat decimalFormat = new DecimalFormat("0.00");
                if (columnDefinition.getValueFormatPattern() != null) {
                    decimalFormat = new DecimalFormat(columnDefinition.getValueFormatPattern());
                }

                value = new BigDecimal(decimalFormat.format(dataCell.getNumericCellValue()));
            } else if (valueField.getType().isAssignableFrom(Integer.class)) {
                value = Double.valueOf(dataCell.getNumericCellValue()).intValue();
            } else if (valueField.getType().isAssignableFrom(String.class)) {
                value = dataCell.getStringCellValue();
            }

            PropertyUtils.setProperty(data, valueField.getName(), value);
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }
    }

    /**
     * 转换Map格式数据为确定的自定义类型数据，要求Map中的Key尽可能多的与自定义类型Field名称一致
     *
     * @param sampleCustomizedObject 自定义类型示例，可预先设置一些通用Field的值
     * @param dataSetInMap           Map格式的数据集
     * @param <T>                    自定义数据类型
     * @return 自定义类型数据集
     * @throws IllegalArgumentException 发生错误时返回无效参数异常
     */
    public static <T> List<T> transformMapToCustomizedType(T sampleCustomizedObject, List<Object> dataSetInMap)
            throws IllegalArgumentException {
        List<T> dataSetOfCustomizedType = new ArrayList<T>();

        try {

            for (Object object : dataSetInMap) {
                Map objectInMap = (Map) object;
                T customizedObject = (T) BeanUtils.cloneBean(sampleCustomizedObject);

                Field[] fields = sampleCustomizedObject.getClass().getDeclaredFields();
                for (Field field : fields) {
                    PropertyUtils.setProperty(customizedObject, field.getName(), objectInMap.get(field.getName()));
                }

                dataSetOfCustomizedType.add(customizedObject);
            }
        } catch (Exception ex) {
            throw new IllegalArgumentException(ex);
        }

        return dataSetOfCustomizedType;
    }

}
