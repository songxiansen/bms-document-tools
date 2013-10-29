package bc.bms.common.workbook;

import bc.bms.common.workbook.model.CalculateColumnDefinition;
import bc.bms.common.workbook.model.ColumnDefinition;
import bc.bms.common.workbook.model.HeaderCellDefinition;
import bc.bms.common.workbook.model.Item;
import bc.bms.common.workbook.model.ItemData;
import bc.bms.common.workbook.model.Project;
import bc.bms.common.workbook.model.ProjectStatus;
import bc.bms.common.workbook.model.WorksheetDefinition;
import bc.bms.common.workbook.model.WorksheetPreference;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class WorkbookToolsTest {

    private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");

    @Test
    //@Ignore
    public void exportSimpleWorkbook() throws IOException {
        printInfo("Begin to export simple workbook.");

        WorksheetDefinition worksheetDefinition = new WorksheetDefinition();

        worksheetDefinition.setName("简单Sheet");

        List<HeaderCellDefinition> headerCellDefinitions = new ArrayList<HeaderCellDefinition>();
        headerCellDefinitions.add(new HeaderCellDefinition("表名id", 0, 0, 0, 0, null));
		headerCellDefinitions.add(new HeaderCellDefinition("快捷支付交易流水历史表", 0, 0, 1, 4, null));
		headerCellDefinitions.add(new HeaderCellDefinition("交易总预览", 0, 0, 5, 7, null));
		headerCellDefinitions.add(new HeaderCellDefinition("交易总金额", 0, 0, 8, 10, null));
		headerCellDefinitions.add(new HeaderCellDefinition("类型", 1, 1, 0, 0, null));
		headerCellDefinitions.add(new HeaderCellDefinition("长度", 1, 1, 1, 1, null));
		headerCellDefinitions.add(new HeaderCellDefinition("空值", 1, 1, 2, 2, null));
		headerCellDefinitions.add(new HeaderCellDefinition("缺省值", 1, 1, 3, 3, null));
		headerCellDefinitions.add(new HeaderCellDefinition("中文名称", 1, 1, 4, 4, null));
		headerCellDefinitions.add(new HeaderCellDefinition("1", 1, 1, 5, 5, null));
		headerCellDefinitions.add(new HeaderCellDefinition("2", 1, 1, 6, 6, null));
		headerCellDefinitions.add(new HeaderCellDefinition("3", 1, 1, 7, 7, null));
		headerCellDefinitions.add(new HeaderCellDefinition("1", 1, 1, 8, 8, null));
		headerCellDefinitions.add(new HeaderCellDefinition("2", 1, 1, 9, 9, null));
		headerCellDefinitions.add(new HeaderCellDefinition("3", 1, 1, 10, 10, null));

        
        worksheetDefinition.setHeaderCellDefinitions(headerCellDefinitions);

//        Map<Integer, ColumnDefinition>
//                columnDefinitions = new HashMap<Integer, ColumnDefinition>();
//
//        ColumnDefinition idColumnDefinition = new ColumnDefinition();
//        idColumnDefinition.setFieldNameOfValue("id");
//        idColumnDefinition.setSampleColumnModel(new Item(null, null));
//        idColumnDefinition.setReadOnlyCellColor(ColorPicker.getColor("#F6F6F6"));
//        idColumnDefinition.setHidden(false);
//        columnDefinitions.put(0, idColumnDefinition);
//
//        ColumnDefinition nameColumnDefinition = new ColumnDefinition();
//        nameColumnDefinition.setFieldNameOfValue("name");
//        nameColumnDefinition.setSampleColumnModel(new Item(null, null));
//        nameColumnDefinition.setReadOnlyCellColor(ColorPicker.getColor("#F6F6F6"));
//        nameColumnDefinition.setWidth(200);
//        nameColumnDefinition.setDataColumnIndex(0);
//        columnDefinitions.put(1, nameColumnDefinition);
        
      
			Map<Integer, ColumnDefinition>
			columnDefinitions = new HashMap<Integer, ColumnDefinition>();
			
			int index = 0;
			for (Field field : Project.class.getDeclaredFields()) {
			    ColumnDefinition columnDefinition = new ColumnDefinition();
			
			    columnDefinition.setFieldNameOfValue(field.getName());
			    columnDefinition.setSampleColumnModel(new Project());
			
			    if (index != 0) {
			        columnDefinition.setDataColumnIndex(0);
			    }
			
			    columnDefinition.setEditable((index == 0) ? false : true);
			    columnDefinition.setEditableCellColor(ColorPicker.getColor("#E2FF9E"));
			    columnDefinition.setReadOnlyCellColor(ColorPicker.getColor("#F3F3F3"));
			    columnDefinitions.put(index, columnDefinition);
			
			    index++;
			}
			worksheetDefinition.setColumnDefinitions(columnDefinitions);
			
			Map<Integer, List>  dataSetOfColumns = new HashMap<Integer, List>();
			for (int i = 0; i < 30; i++) {
			    Project project = new Project(
			            i,(i*1.5 + 1)+"cm",  (i*2 + 1)+"P",
			            new BigDecimal(1 + i + 0.3), new BigDecimal(1 + i + 0.6*i),
			            ProjectStatus.getProjectStatus(i+1).getCode());
			
			    WorksheetDefinition.addDataOfColumn(
			            dataSetOfColumns, WorksheetPreference.WORKSHEET_PRIMARY_COLUMN_INDEX, project);
			}
      
       
        WorksheetPreference worksheetPreference = new WorksheetPreference();
        worksheetPreference.setImportable(true);
        worksheetPreference.setHeaderRowHeight((short) 25);
        worksheetPreference.setRowHeight((short) 30);
        worksheetDefinition.setWorksheetPreference(worksheetPreference);
        
        XSSFWorkbook workbook = WorkbookWriter.writeToWorkbook(null, worksheetDefinition, dataSetOfColumns);
        WorkbookWriter.appendDataToWorksheet(workbook.getSheetAt(0), worksheetDefinition, dataSetOfColumns);
        WorkbookWriter.appendDataToWorksheet(workbook.getSheetAt(0), worksheetDefinition, dataSetOfColumns);
        WorkbookWriter.appendDataToWorksheet(workbook.getSheetAt(0), worksheetDefinition, dataSetOfColumns);
        WorkbookWriter.appendDataToWorksheet(workbook.getSheetAt(0), worksheetDefinition, dataSetOfColumns);
        File file = new File("G:\\bms-document-tools-master\\tmp\\SampleWorkbook.xlsx");
        if (!file.exists()) {
            file.createNewFile();
        }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        printInfo("Export simple workbook end.");
        System.out.println();
    }

    @Test
    //@Ignore
    public void readSimpleWorkbook() throws FileNotFoundException {
        printInfo("Begin to read simple workbook.");

        File file = new File("G:\\bms-document-tools-master\\tmp\\SampleWorkbook.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        Map<Integer, List<Object>>
                dataSetSeparatedByType = WorkbookReader.collectDataFromWorkbook(fileInputStream);

        for (Map.Entry<Integer, List<Object>>
                dataSetSeparatedByTypeEntry : dataSetSeparatedByType.entrySet()) {
            Integer columnIndex = dataSetSeparatedByTypeEntry.getKey();
            List<Object> dataSet = dataSetSeparatedByTypeEntry.getValue();

            DecimalFormat decimalFormat = new DecimalFormat("0000");
            for (Object data : dataSet) {
                printInfo("[" + decimalFormat.format(columnIndex) + "] " + data.toString());
            }
        }

        printInfo("Read simple workbook end.");
        System.out.println();
    }

    @Test
    //@Ignore
    public void exportProjects() throws IOException {
        printInfo("Begin to export projects.");

        WorksheetDefinition worksheetDefinition = new WorksheetDefinition();

        worksheetDefinition.setName("交易记录");

        List<HeaderCellDefinition> headerCellDefinitions = new ArrayList<HeaderCellDefinition>();
        headerCellDefinitions.add(new HeaderCellDefinition("交易Id", 0, 0, 0, 0, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易人姓名", 0, 0, 1, 1, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易时间", 0, 0, 2, 2, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易金额", 0, 0, 3, 3, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易类型", 0, 0, 4, 4, null));
		headerCellDefinitions.add(new HeaderCellDefinition("交易结果", 0, 0, 5, 5, null));


        worksheetDefinition.setHeaderCellDefinitions(headerCellDefinitions);

        Map<Integer, ColumnDefinition>
                columnDefinitions = new HashMap<Integer, ColumnDefinition>();

        int index = 0;
        for (Field field : Project.class.getDeclaredFields()) {
            ColumnDefinition columnDefinition = new ColumnDefinition();

            columnDefinition.setFieldNameOfValue(field.getName());
            columnDefinition.setSampleColumnModel(new Project());

            if (index != 0) {
                columnDefinition.setDataColumnIndex(0);
            }

            columnDefinition.setEditable((index == 0) ? false : true);
            columnDefinition.setEditableCellColor(ColorPicker.getColor("#E2FF9E"));
            columnDefinition.setReadOnlyCellColor(ColorPicker.getColor("#F3F3F3"));
            columnDefinitions.put(index, columnDefinition);

            index++;
        }
        worksheetDefinition.setColumnDefinitions(columnDefinitions);

        Map<Integer, List>  dataSetOfColumns = new HashMap<Integer, List>();
        for (int i = 0; i < 30; i++) {
            Project project = new Project(
                    i, "NO." + (i*1.5 + 1), "P0" + (i*2 + 1),
                    new BigDecimal(1 + i + 0.3), new BigDecimal(1 + i + 0.6*i),
                    ProjectStatus.getProjectStatus(i+1).getName());

            WorksheetDefinition.addDataOfColumn(
                    dataSetOfColumns, WorksheetPreference.WORKSHEET_PRIMARY_COLUMN_INDEX, project);
        }

        WorksheetPreference worksheetPreference = new WorksheetPreference();
        worksheetPreference.setImportable(true);
        worksheetPreference.setHeaderRowHeight((short) 25);
        worksheetPreference.setRowHeight((short) 25);
        worksheetDefinition.setWorksheetPreference(worksheetPreference);

        XSSFWorkbook workbook = WorkbookWriter.writeToWorkbook(null, worksheetDefinition, dataSetOfColumns);
        File file = new File("G:\\bms-document-tools-master\\tmp\\ProjectWorkbook.xlsx");
        if (!file.exists()) {
            file.createNewFile();
        }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        printInfo("Export projects end.");
        System.out.println();
    }

    @Test
    //@Ignore
    public void readProjects() throws FileNotFoundException {
        printInfo("Begin to read projects.");

        File file = new File("G:\\bms-document-tools-master\\tmp\\ProjectWorkbook.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        Map<Integer, List<Object>>
                dataSetSeparatedByType = WorkbookReader.collectDataFromWorkbook(fileInputStream);

        for (Map.Entry<Integer, List<Object>>
                dataSetSeparatedByTypeEntry : dataSetSeparatedByType.entrySet()) {
            Integer columnIndex = dataSetSeparatedByTypeEntry.getKey();
            List<Object> dataSet = dataSetSeparatedByTypeEntry.getValue();

            DecimalFormat decimalFormat = new DecimalFormat("0000");
            for (Object data : dataSet) {
                printInfo("[" + decimalFormat.format(columnIndex) + "] " + data.toString());
            }
        }

        printInfo("Read projects end.");
        System.out.println();
    }

    @Test
    //@Ignore
    public void exportProjectsInMap() throws IOException {
        printInfo("Begin to export projects in map.");

        WorksheetDefinition worksheetDefinition = new WorksheetDefinition();

        worksheetDefinition.setName("交易表In Map");

        List<HeaderCellDefinition> headerCellDefinitions = new ArrayList<HeaderCellDefinition>();
        
        headerCellDefinitions.add(new HeaderCellDefinition("交易Id", 0, 0, 0, 0, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易编号", 0, 0, 1, 1, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易数字", 0, 0, 2, 2, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易金额", 0, 0, 3, 3, null));
        headerCellDefinitions.add(new HeaderCellDefinition("金额预算", 0, 0, 4, 4, null));
        headerCellDefinitions.add(new HeaderCellDefinition("交易状态", 0, 0, 5, 5, null));
        headerCellDefinitions.add(new HeaderCellDefinition("快捷支付交易流水表记录", 1, 1, 0, 5, null));
        worksheetDefinition.setHeaderCellDefinitions(headerCellDefinitions);

        Map<Integer, ColumnDefinition>
                columnDefinitions = new HashMap<Integer, ColumnDefinition>();

        int index = 0;
        String[] keys = {"id", "code", "name", "costAmount", "fundAmount", "status"};
        for (String key : keys) {
            ColumnDefinition columnDefinition = new ColumnDefinition();

            columnDefinition.setFieldNameOfValue(key);
            columnDefinition.setSampleColumnModel(new HashMap<String, Object>());

            if (index != 0) {
                columnDefinition.setDataColumnIndex(0);
            }

            switch (index) {
                case 0:
                    columnDefinition.setFieldTypeOfValue(Integer.class.getName());
                    break;
                case 3:
                    columnDefinition.setFieldTypeOfValue(BigDecimal.class.getName());
                    break;
                case 4:
                    columnDefinition.setFieldTypeOfValue(BigDecimal.class.getName());
                    break;
                default:
                    columnDefinition.setFieldTypeOfValue(String.class.getName());
                    break;
            }

            columnDefinition.setEditable((index == 0) ? false : true);
            columnDefinition.setEditableCellColor(ColorPicker.getColor("#E2FF9E"));
            columnDefinition.setReadOnlyCellColor(ColorPicker.getColor("#F3F3F3"));
            columnDefinitions.put(index, columnDefinition);

            index++;
        }
        worksheetDefinition.setColumnDefinitions(columnDefinitions);

        Map<Integer, List>  dataSetOfColumns = new HashMap<Integer, List>();
        for (int i = 0; i < 56; i++) {
            Map<String, Object> project = new HashMap<String, Object>();
            project.put(keys[0], i);
            project.put(keys[1], "No.编号" + (i + 1));
            project.put(keys[2], "交易数" + (i*2 + 0.5*i));
            project.put(keys[3], new BigDecimal(1 + i*1.5 + 0.2));
            project.put(keys[4], new BigDecimal(1 + i*2 + 0.6));
            project.put(keys[5], ProjectStatus.getProjectStatus(i + 1).getName());

            WorksheetDefinition.addDataOfColumn(
                    dataSetOfColumns, WorksheetPreference.WORKSHEET_PRIMARY_COLUMN_INDEX, project);
        }

        WorksheetPreference worksheetPreference = new WorksheetPreference();
        worksheetPreference.setImportable(true);
        worksheetPreference.setHeaderRowHeight((short) 25);
        worksheetPreference.setRowHeight((short) 25);
        worksheetDefinition.setWorksheetPreference(worksheetPreference);

        XSSFWorkbook workbook = WorkbookWriter.writeToWorkbook(null, worksheetDefinition, dataSetOfColumns);
        File file = new File("G:\\bms-document-tools-master\\tmp\\ProjectInMapWorkbook.xlsx");
        if (!file.exists()) {
            file.createNewFile();
        }
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);

        printInfo("Export projects in map end.");
        System.out.println();
    }

    @Test
    //@Ignore
    public void readProjectsInMap() throws FileNotFoundException {
        printInfo("Begin to read projects in map.");

        File file = new File("G:\\bms-document-tools-master\\tmp\\ProjectInMapWorkbook.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        Map<Integer, List<Object>>
                dataSetSeparatedByType = WorkbookReader.collectDataFromWorkbook(fileInputStream);

        for (Map.Entry<Integer, List<Object>>
                dataSetSeparatedByTypeEntry : dataSetSeparatedByType.entrySet()) {
            Integer columnIndex = dataSetSeparatedByTypeEntry.getKey();
            List<Object> dataSet = dataSetSeparatedByTypeEntry.getValue();

            DecimalFormat decimalFormat = new DecimalFormat("0000");
            for (Object data : dataSet) {
                printInfo("[" + decimalFormat.format(columnIndex) + "] " + data.toString());
            }
        }

        printInfo("Read projects in map end.");
        System.out.println();
    }

    @Test
    //@Ignore
    public void readProjectsInMapAndTransformToProjectType() throws FileNotFoundException {
        printInfo("Begin to read projects in map and transform to 'Project'.");

        File file = new File("G:\\bms-document-tools-master\\tmp\\ProjectInMapWorkbook.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        Map<Integer, List<Object>>
                dataSetSeparatedByType = WorkbookReader.collectDataFromWorkbook(fileInputStream);

        for (Map.Entry<Integer, List<Object>>
                dataSetSeparatedByTypeEntry : dataSetSeparatedByType.entrySet()) {
            Integer columnIndex = dataSetSeparatedByTypeEntry.getKey();
            List<Object> dataSet = dataSetSeparatedByTypeEntry.getValue();

            List<Project> projects = WorkbookReader.transformMapToCustomizedType(new Project(), dataSet);

            DecimalFormat decimalFormat = new DecimalFormat("0000");
            for (Project project : projects) {
                printInfo("[" + decimalFormat.format(columnIndex) + "] " + project.toString());
            }
        }

        printInfo("Read projects in map and transform to 'Project' end.");
        System.out.println();
    }

    @Test
    @Ignore
    public void getCalculateColumnDefinitionFromTemplate() {
        CalculateColumnDefinition<Item> sampleCalculateColumnDefinition = new CalculateColumnDefinition<Item>();
        sampleCalculateColumnDefinition.setSampleColumnModel(new Item());
        Map<String, List<CalculateColumnDefinition<Item>>> calculateColumnDefinitionsSeparatedByWorksheet
                = WorkbookReader.getCalculateColumnDefinitionFromTemplate(
                "G:\\bms-document-tools-master\\tmp\\CalculateTemplate.xlsx", sampleCalculateColumnDefinition, new String[]{"id", "name"});

        if (calculateColumnDefinitionsSeparatedByWorksheet != null) {
            for (Map.Entry<String, List<CalculateColumnDefinition<Item>>>
                    calculateColumnDefinitionsSeparatedByWorksheetEntry
                    : calculateColumnDefinitionsSeparatedByWorksheet.entrySet()) {
                String worksheetName = calculateColumnDefinitionsSeparatedByWorksheetEntry.getKey();
                List<CalculateColumnDefinition<Item>> calculateColumnDefinitions
                        = calculateColumnDefinitionsSeparatedByWorksheetEntry.getValue();

                printInfo("Worksheet: " + worksheetName);

                if (calculateColumnDefinitions != null && calculateColumnDefinitions.size() > 0) {
                    for (CalculateColumnDefinition<Item> calculateColumnDefinition : calculateColumnDefinitions) {
                        printInfo(calculateColumnDefinition.toString());

                        if (calculateColumnDefinition.getSampleColumnModel() != null) {
                            printInfo("\t> " + calculateColumnDefinition.getSampleColumnModel());
                        }
                    }
                } else {
                    printInfo("> not found any data.");
                }

                printInfo(null);
            }
        }
    }

    private void printInfo(String information) {
        if (information != null && information.length() > 0) {
            System.out.println("[" + sdf.format(new Date()) + "] " + information);
        } else {
            System.out.println();
        }
    }

    @Test
    public void getDataTimeInString() {
        SimpleDateFormat dateTimeFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Calendar calendar = Calendar.getInstance();
        calendar.set(2014, 1, 1, 0, 0, 0);

        if (true) {
            calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMaximum(Calendar.DAY_OF_MONTH));
        } else {
            calendar.set(Calendar.DAY_OF_MONTH, calendar.getActualMinimum(Calendar.DAY_OF_MONTH));
        }

        System.out.println(dateTimeFormat.format(calendar.getTime()));
    }

}
