package com.feiyizhan.excel.export.utils;

import com.feiyizhan.excel.export.utils.data.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.jupiter.api.Test;

import java.util.*;

import static com.feiyizhan.excel.export.utils.data.ExcelCellValueTypeData.MULTI_LINE_TEXT;
import static com.feiyizhan.excel.export.utils.data.ExcelCellValueTypeData.SINGLE_LINE_TEXT;

/**
 * Excel导出工具测试类
 * @author 徐明龙 XuMingLong 2023-03-06
 */
public class ExportExcelUtilTests {


    /**
     * 测试简单根据模版生成Excel文件 (一个Sheet模版，一个Sheet输出，一个表格）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_1(){
        String templateFile = "template.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_1.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        //设置Sheet数据
        ExcelSheetData sheetData = new ExcelSheetData();
        sheetDataList.add(sheetData);
        sheetData.setTemplateSheetIndex(0);
        sheetData.setOutSheetName("测试Sheet1");
        sheetData.setEmptied(false);
        //设置表格数据
        List<ExcelTableData> tableDataList = new ArrayList<>();
        sheetData.setTableDataList(tableDataList);
        ExcelTableData tableData = new ExcelTableData();
        tableDataList.add(tableData);
        //填充表格数据
        fillTableData(tableData,2,2,10,18);
        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }


    /**
     * 测试简单根据模版生成Excel文件  (一个Sheet模版，一个Sheet输出，多个表格）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_2(){
        String templateFile = "template2.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_2.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        //设置Sheet数据
        ExcelSheetData sheetData = new ExcelSheetData();
        sheetDataList.add(sheetData);
        sheetData.setTemplateSheetIndex(0);
        sheetData.setOutSheetName("测试Sheet1");
        sheetData.setEmptied(false);
        //设置表格数据
        List<ExcelTableData> tableDataList = new ArrayList<>();
        sheetData.setTableDataList(tableDataList);
        ExcelTableData tableData = new ExcelTableData();
        tableDataList.add(tableData);
        //填充表格1数据
        fillTableData(tableData,3,3,10,18);

        ExcelTableData tableData2 = new ExcelTableData();
        tableDataList.add(tableData2);
        //填充表格2数据
        fillTableData(tableData2,17+9,17+9,5,18);

        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }



    /**
     * 测试简单根据模版生成Excel文件  (一个Sheet模版，多个Sheet输出，多个表格）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_3(){
        String templateFile = "template3.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_3.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        for(int i=1;i<4;i++){
            //设置Sheet数据
            ExcelSheetData sheetData = new ExcelSheetData();
            sheetDataList.add(sheetData);
            sheetData.setTemplateSheetIndex(0);
            sheetData.setOutSheetName("测试Sheet"+i);
            sheetData.setEmptied(false);
            //设置表格数据
            List<ExcelTableData> tableDataList = new ArrayList<>();
            sheetData.setTableDataList(tableDataList);
            ExcelTableData tableData = new ExcelTableData();
            tableDataList.add(tableData);
            //填充表格1数据
            fillTableData(tableData,3,3,10,18);

            ExcelTableData tableData2 = new ExcelTableData();
            tableDataList.add(tableData2);
            //填充表格2数据
            fillTableData(tableData2,17+9,17+9,5,18);
        }


        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }


    /**
     * 测试简单根据模版生成Excel文件  (多个Sheet模版，混合）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_4(){
        String templateFile = "template4.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_4.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        for(int i=1;i<4;i++){
            //设置Sheet数据
            ExcelSheetData sheetData = new ExcelSheetData();
            sheetDataList.add(sheetData);
            sheetData.setTemplateSheetIndex(0);
            sheetData.setOutSheetName("测试Sheet模版1结果"+i);
            sheetData.setEmptied(false);
            //设置表格数据
            List<ExcelTableData> tableDataList = new ArrayList<>();
            sheetData.setTableDataList(tableDataList);
            ExcelTableData tableData = new ExcelTableData();
            tableDataList.add(tableData);
            //填充表格1数据
            fillTableData(tableData,3,3,10,18);

            ExcelTableData tableData2 = new ExcelTableData();
            tableDataList.add(tableData2);
            //填充表格2数据
            fillTableData(tableData2,17+9,17+9,5,18);
        }

        for(int i=1;i<3;i++){
            //设置Sheet数据
            ExcelSheetData sheetData = new ExcelSheetData();
            sheetDataList.add(sheetData);
            sheetData.setTemplateSheetIndex(1);
            sheetData.setOutSheetName("测试Sheet模版2结果"+i);
            sheetData.setEmptied(false);
            //设置表格数据
            List<ExcelTableData> tableDataList = new ArrayList<>();
            sheetData.setTableDataList(tableDataList);
            ExcelTableData tableData = new ExcelTableData();
            tableDataList.add(tableData);
            //填充表格1数据
            fillTableData(tableData,3,3,10,18);
        }

        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }


    /**
     * 测试简单根据模版生成Excel文件  (一个Sheet模版，一个Sheet输出，多个表格,参考行不在表格内）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_5(){
        String templateFile = "template5.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_5.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        //设置Sheet数据
        ExcelSheetData sheetData = new ExcelSheetData();
        sheetDataList.add(sheetData);
        sheetData.setTemplateSheetIndex(0);
        sheetData.setOutSheetName("测试Sheet1");
        sheetData.setEmptied(false);
        //设置表格数据
        List<ExcelTableData> tableDataList = new ArrayList<>();
        sheetData.setTableDataList(tableDataList);
        ExcelTableData tableData = new ExcelTableData();
        tableDataList.add(tableData);
        //填充表格1数据
        fillTableData(tableData,3,17+9,10,18);

        ExcelTableData tableData2 = new ExcelTableData();
        tableDataList.add(tableData2);
        //填充表格2数据
        fillTableData(tableData2,17+9,3,5,18);

        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }

    /**
     * 测试简单根据模版生成Excel文件  (一个Sheet模版，简单的单表格没有自定义样式的输出）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_6(){
        String templateFile = "template6.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_6.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        //设置Sheet数据
        ExcelSheetData sheetData = new ExcelSheetData();
        sheetDataList.add(sheetData);
        sheetData.setTemplateSheetIndex(0);
        sheetData.setOutSheetName("测试Sheet1");
        sheetData.setEmptied(false);
        //设置表格数据
        List<ExcelTableData> tableDataList = new ArrayList<>();
        sheetData.setTableDataList(tableDataList);
        ExcelTableData tableData = new ExcelTableData();
        tableDataList.add(tableData);
        //填充表格1数据
        fillTableData(tableData,3,10);

        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }


    /**
     * 测试简单根据模版生成Excel文件 (测试单元格超过64000的问题）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_7(){
        String templateFile = "template7.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_7.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        //设置Sheet数据
        ExcelSheetData sheetData = new ExcelSheetData();
        sheetDataList.add(sheetData);
        sheetData.setTemplateSheetIndex(0);
        sheetData.setOutSheetName("测试Sheet1");
        sheetData.setEmptied(false);
        //设置表格数据
        List<ExcelTableData> tableDataList = new ArrayList<>();
        sheetData.setTableDataList(tableDataList);
        ExcelTableData tableData = new ExcelTableData();
        tableDataList.add(tableData);
        //填充表格数据
        fillTableData(tableData,3,3,70000,18);
        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }


    /**
     * 测试简单根据模版生成Excel文件 (解决测试单元格超过64000的问题）
     * @author 徐明龙 XuMingLong 2023-03-06
     * @return void
     */
    @Test
    public void test_generateAndExportExcelToFile_8(){
        String templateFile = "template8.xlsx";
        String outFile = "D:\\test_generateAndExportExcelToFile_8.xlsx";
        ExcelFileData excelFileData = new ExcelFileData();
        //导出的文件名
        excelFileData.setOutFileName(outFile);
        //导出的Sheet数据
        List<ExcelSheetData> sheetDataList = new ArrayList<>();
        excelFileData.setSheetDataList(sheetDataList);
        //导出的变量数据
        Map<String, ExcelCellData> variableMap = new HashMap<>();
        excelFileData.setVariableMap(variableMap);
        //设置Sheet数据
        ExcelSheetData sheetData = new ExcelSheetData();
        sheetDataList.add(sheetData);
        sheetData.setTemplateSheetIndex(0);
        sheetData.setOutSheetName("测试Sheet1");
        sheetData.setEmptied(false);
        //设置表格数据
        List<ExcelTableData> tableDataList = new ArrayList<>();
        sheetData.setTableDataList(tableDataList);
        ExcelTableData tableData = new ExcelTableData();
        tableDataList.add(tableData);
        //填充表格数据
        fillTableDataForFixMaximumCellStylesIssue(tableData,3,3,70000,18);
        //填充模版变量Map
        fillVariableMap(variableMap);
        //生成并导出Excel文件
        ExportExcelUtil.generateAndExportExcelToFile(excelFileData,FileUtil.getFileInputStream(templateFile));
    }
    /**
     * 填充表格数据
     * @author 徐明龙 XuMingLong 2023-03-08
     * @param tableData
     * @param beginRowNo  起始行号
     * @param rowSize  行数
     * @return void
     */
    private void fillTableData(ExcelTableData tableData,int beginRowNo,int rowSize){
        tableData.setBeginRowNo(beginRowNo);
        //设置行数据
        List<ExcelRowData> rowDataList = new ArrayList<>();
        tableData.setRowDataList(rowDataList);
        for(int i=0;i<rowSize;i++){
            ExcelRowData rowData = new ExcelRowData();
            rowDataList.add(rowData);
            List<ExcelCellData> cellDataList = new ArrayList<>();
            rowData.setCellDataList(cellDataList);
            //序号
            cellDataList.add(ExcelCellData.builder().value(i+1).build());
            //编号
            cellDataList.add(ExcelCellData.builder().value("BH0001").build());
            //客户名称
            cellDataList.add(ExcelCellData.builder().value("客户A").build());
            //统一社会信用代码
            cellDataList.add(ExcelCellData.builder().value("3343241231231312").build());
            //客户信息备注
            cellDataList.add(ExcelCellData.builder().value(Arrays.asList("备注1","备注2","备注3")).build());
            //联系人
            cellDataList.add(ExcelCellData.builder().value("旺旺").build());
            //手机号
            cellDataList.add(ExcelCellData.builder().value("13111111111").build());
            //销售负责人
            cellDataList.add(ExcelCellData.builder().value("喵喵").build());
            //项目负责人
            cellDataList.add(ExcelCellData.builder().value("唧唧").build());
            //订单编号
            cellDataList.add(ExcelCellData.builder().value("DDBH000001").build());
            //合同编号
            cellDataList.add(ExcelCellData.builder().value("HTBH000001").build());
            //订单类型
            cellDataList.add(ExcelCellData.builder().value("套餐").build());
            //订单状态
            cellDataList.add(ExcelCellData.builder().value("已回款").build());
            //开始日期
            cellDataList.add(ExcelCellData.builder().value("2023-03-04").build());
            //结束日期
            cellDataList.add(ExcelCellData.builder().value("2023-05-04").build());
            //有效期/天
            cellDataList.add(ExcelCellData.builder().value(61).build());
            //套餐时长
            cellDataList.add(ExcelCellData.builder().value(1000.5).build());
            //订单备注
            cellDataList.add(ExcelCellData.builder().value("").build());
        }
    }

    /**
     * 填充表格数据(修复超过最大单元格样式的问题）
     * @author 徐明龙 XuMingLong 2023-03-08
     * @param tableData
     * @param beginRowNo  起始行号
     * @param referRowNo  参考样式的行号
     * @param rowSize  行数
     * @param columnSize  列数
     * @return void
     */
    private void fillTableDataForFixMaximumCellStylesIssue(ExcelTableData tableData,int beginRowNo,int referRowNo,int rowSize,int columnSize){
        tableData.setBeginRowNo(beginRowNo);
        //设置行数据
        List<ExcelRowData> rowDataList = new ArrayList<>();
        tableData.setRowDataList(rowDataList);
        //创建奇数行的样式
        final List<ExcelCellCustomData> oddCellCustomDataList = createOddRowStyle(referRowNo,columnSize);
        //创建偶数数行的样式
        final List<ExcelCellCustomData> evenCellCustomDataList = createEvenRowStyle(referRowNo,columnSize);
        //创建引用单元格样式
        final List<ExcelCellCustomData> oddReferCellCustomDataList = createReferRowStyle(beginRowNo+1,columnSize);
        final List<ExcelCellCustomData> evenReferCellCustomDataList = createReferRowStyle(beginRowNo,columnSize);
        for(int i=0;i<rowSize;i++){
            ExcelRowData rowData = new ExcelRowData();
            rowDataList.add(rowData);
            List<ExcelCellData> cellDataList = new ArrayList<>();
            rowData.setCellDataList(cellDataList);

            List<ExcelCellCustomData> cellCustomDataList = null;
            if(i<2){
                cellCustomDataList = i%2==0?evenCellCustomDataList:oddCellCustomDataList;
            }else{
                cellCustomDataList = i%2==0?evenReferCellCustomDataList:oddReferCellCustomDataList;

            }
            int column = 0;
            //序号
            cellDataList.add(ExcelCellData.builder().value(i+1).cellCustomData(cellCustomDataList.get(column++)).build());
            //编号
            cellDataList.add(ExcelCellData.builder().value("BH0001").cellCustomData(cellCustomDataList.get(column++)).build());
            //客户名称
            cellDataList.add(ExcelCellData.builder().value("客户A").cellCustomData(cellCustomDataList.get(column++)).build());
            //统一社会信用代码
            cellDataList.add(ExcelCellData.builder().value("3343241231231312")
                .cellCustomData(cellCustomDataList.get(column++)).build());
            //客户信息备注
            cellDataList.add(ExcelCellData.builder().value(Arrays.asList("备注1","备注2","备注3"))
                .cellCustomData(cellCustomDataList.get(column++)).build());
            //联系人
            cellDataList.add(ExcelCellData.builder().value("旺旺").cellCustomData(cellCustomDataList.get(column++)).build());
            //手机号
            cellDataList.add(ExcelCellData.builder().value("13111111111").cellCustomData(cellCustomDataList.get(column++)).build());
            //销售负责人
            cellDataList.add(ExcelCellData.builder().value("喵喵").cellCustomData(cellCustomDataList.get(column++)).build());
            //项目负责人
            cellDataList.add(ExcelCellData.builder().value("唧唧").cellCustomData(cellCustomDataList.get(column++)).build());
            //订单编号
            cellDataList.add(ExcelCellData.builder().value("DDBH000001").cellCustomData(cellCustomDataList.get(column++)).build());
            //合同编号
            cellDataList.add(ExcelCellData.builder().value("HTBH000001").cellCustomData(cellCustomDataList.get(column++)).build());
            //订单类型
            cellDataList.add(ExcelCellData.builder().value("套餐").cellCustomData(cellCustomDataList.get(column++)).build());
            //订单状态
            cellDataList.add(ExcelCellData.builder().value("已回款").cellCustomData(cellCustomDataList.get(column++)).build());
            //开始日期
            cellDataList.add(ExcelCellData.builder().value("2023-03-04").cellCustomData(cellCustomDataList.get(column++)).build());
            //结束日期
            cellDataList.add(ExcelCellData.builder().value("2023-05-04").cellCustomData(cellCustomDataList.get(column++)).build());
            //有效期/天
            cellDataList.add(ExcelCellData.builder().value(61).cellCustomData(cellCustomDataList.get(column++)).build());
            //套餐时长
            cellDataList.add(ExcelCellData.builder().value(1000.5).cellCustomData(cellCustomDataList.get(column++)).build());
            //订单备注
            cellDataList.add(ExcelCellData.builder().value("").cellCustomData(cellCustomDataList.get(column++)).build());
        }
    }


    /**
     * 填充表格数据
     * @author 徐明龙 XuMingLong 2023-03-08
     * @param tableData
     * @param beginRowNo  起始行号
     * @param referRowNo  参考样式的行号
     * @param rowSize  行数
     * @param columnSize  列数
     * @return void
     */
    private void fillTableData(ExcelTableData tableData,int beginRowNo,int referRowNo,int rowSize,int columnSize){
        tableData.setBeginRowNo(beginRowNo);
        //设置行数据
        List<ExcelRowData> rowDataList = new ArrayList<>();
        tableData.setRowDataList(rowDataList);
        //创建奇数行的样式
        final List<ExcelCellCustomData> oddCellCustomDataList = createOddRowStyle(referRowNo,columnSize);
        //创建偶数数行的样式
        final List<ExcelCellCustomData> evenCellCustomDataList = createEvenRowStyle(referRowNo,columnSize);
        for(int i=0;i<rowSize;i++){
            ExcelRowData rowData = new ExcelRowData();
            rowDataList.add(rowData);
            List<ExcelCellData> cellDataList = new ArrayList<>();
            rowData.setCellDataList(cellDataList);
            List<ExcelCellCustomData> cellCustomDataList = i%2==0?evenCellCustomDataList:oddCellCustomDataList;
            int column = 0;
            //序号
            cellDataList.add(ExcelCellData.builder().value(i+1).cellCustomData(cellCustomDataList.get(column++)).build());
            //编号
            cellDataList.add(ExcelCellData.builder().value("BH0001").cellCustomData(cellCustomDataList.get(column++)).build());
            //客户名称
            cellDataList.add(ExcelCellData.builder().value("客户A").cellCustomData(cellCustomDataList.get(column++)).build());
            //统一社会信用代码
            cellDataList.add(ExcelCellData.builder().value("3343241231231312")
                .cellCustomData(cellCustomDataList.get(column++)).build());
            //客户信息备注
            cellDataList.add(ExcelCellData.builder().value(Arrays.asList("备注1","备注2","备注3"))
                .cellCustomData(cellCustomDataList.get(column++)).build());
            //联系人
            cellDataList.add(ExcelCellData.builder().value("旺旺").cellCustomData(cellCustomDataList.get(column++)).build());
            //手机号
            cellDataList.add(ExcelCellData.builder().value("13111111111").cellCustomData(cellCustomDataList.get(column++)).build());
            //销售负责人
            cellDataList.add(ExcelCellData.builder().value("喵喵").cellCustomData(cellCustomDataList.get(column++)).build());
            //项目负责人
            cellDataList.add(ExcelCellData.builder().value("唧唧").cellCustomData(cellCustomDataList.get(column++)).build());
            //订单编号
            cellDataList.add(ExcelCellData.builder().value("DDBH000001").cellCustomData(cellCustomDataList.get(column++)).build());
            //合同编号
            cellDataList.add(ExcelCellData.builder().value("HTBH000001").cellCustomData(cellCustomDataList.get(column++)).build());
            //订单类型
            cellDataList.add(ExcelCellData.builder().value("套餐").cellCustomData(cellCustomDataList.get(column++)).build());
            //订单状态
            cellDataList.add(ExcelCellData.builder().value("已回款").cellCustomData(cellCustomDataList.get(column++)).build());
            //开始日期
            cellDataList.add(ExcelCellData.builder().value("2023-03-04").cellCustomData(cellCustomDataList.get(column++)).build());
            //结束日期
            cellDataList.add(ExcelCellData.builder().value("2023-05-04").cellCustomData(cellCustomDataList.get(column++)).build());
            //有效期/天
            cellDataList.add(ExcelCellData.builder().value(61).cellCustomData(cellCustomDataList.get(column++)).build());
            //套餐时长
            cellDataList.add(ExcelCellData.builder().value(1000.5).cellCustomData(cellCustomDataList.get(column++)).build());
            //订单备注
            cellDataList.add(ExcelCellData.builder().value("").cellCustomData(cellCustomDataList.get(column++)).build());
        }
    }


    /**
     * 创建参考引用的样式
     * @author 徐明龙 XuMingLong 2023-03-08
     * @param referRowNo
     * @param maxColumnSize
     * @return java.util.List<com.feiyizhan.excel.export.utils.data.ExcelCellCustomData>
     */
    private List<ExcelCellCustomData> createReferRowStyle(int referRowNo, int maxColumnSize) {
        List<ExcelCellCustomData> cellCustomDataList = new ArrayList<>(maxColumnSize);
        for(int i=0;i<maxColumnSize;i++){
            //所有列的样式 （参考行的样式)
            cellCustomDataList.add(ExcelReferCellStyleData.builder()
                .referStyleCell(ExcelCellLocation.builder().row(referRowNo).column(i).build())
                .build());
        }
        return cellCustomDataList;
    }


    /**
     * 创建奇数行的样式
     * @author 徐明龙 XuMingLong 2023-03-07
     * @param referRowNo 参考样式的行
     * @param maxColumnSize
     * @return java.util.List<com.feiyizhan.excel.export.utils.data.ExcelCellCustomData>
     */
    private List<ExcelCellCustomData> createOddRowStyle(int referRowNo,int maxColumnSize){
        ExcelCellFormatter customCellFormatter =(cell)->{
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cell.getSheet().getWorkbook().createFont();
            font.setBold(true);//粗体显示
            font.setColor(IndexedColors.RED.getIndex());
            font.setFontName("微软雅黑");
            font.setFontHeightInPoints((short) 12);
            cellStyle.setFont(font);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置实心填充
            cellStyle.setFillForegroundColor(IndexedColors.BLUE1.index); //设置填充的背景色
        };

        ExcelCellFormatter customCellFormatter2 =(cell)->{
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cell.getSheet().getWorkbook().createFont();
            font.setBold(false);//粗体显示
            font.setColor(IndexedColors.YELLOW.getIndex());
            font.setFontName("微软雅黑");
            font.setFontHeightInPoints((short) 12);
            cellStyle.setFont(font);
        };

        ExcelCellFormatter customCellFormatter3 =(cell)->{
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cell.getSheet().getWorkbook().createFont();
            font.setBold(false);//粗体显示
            font.setColor(IndexedColors.PINK.getIndex());
            font.setFontName("微软雅黑");
            font.setFontHeightInPoints((short) 12);
            cellStyle.setFont(font);
        };
        ExcelCellFormatter customCellFormatter4 =(cell)->{
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cell.getSheet().getWorkbook().createFont();
            font.setBold(false);//粗体显示
            font.setColor(IndexedColors.BLUE.getIndex());
            font.setFontName("微软雅黑");
            font.setFontHeightInPoints((short) 12);
            cellStyle.setFont(font);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置实心填充
            cellStyle.setFillForegroundColor(IndexedColors.YELLOW1.index); //设置填充的背景色
        };

        List<ExcelCellCustomData> cellCustomDataList = new ArrayList<>(maxColumnSize);
        int i = 0;
        //序号列的样式(参考行的样式）
        cellCustomDataList.add(ExcelReferCellStyleData.builder()
            .referStyleCell(ExcelCellLocation.builder().row(referRowNo).column(i).build())
            .build());
        i++;
        //客户编号的样式 (参考行的样式，但创建新不共享的样式)
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .referStyleCell(ExcelCellLocation.builder().row(referRowNo).column(i).build())
            .customFormatterList(Arrays.asList(customCellFormatter))
            .build());
        i++;
        //客户名称的样式 （不参考行的样式，创建不共享的样式）
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .customFormatterList(Arrays.asList(customCellFormatter2))
            .build());
        i++;
        //统一社会信用代码的样式 （参考行的样式，并设置值的类型为单行文本）
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .customFormatterList(Arrays.asList(customCellFormatter3))
            .customCellValueFun(SINGLE_LINE_TEXT.getSetValueFun())
            .build());
        i++;
        //客户的备注的样式 （参考行的样式，并设置值的类型为多行文本）
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .customFormatterList(Arrays.asList(customCellFormatter4,MULTI_LINE_TEXT.getCellFormatter()))
            .customCellValueFun(MULTI_LINE_TEXT.getSetValueFun())
            .build());
        i++;
        for(;i<maxColumnSize;i++){
            //之后的所有列的样式 （参考行的样式)
            cellCustomDataList.add(ExcelReferCellStyleData.builder()
                .referStyleCell(ExcelCellLocation.builder().row(referRowNo).column(i).build())
                .build());
        }
        return cellCustomDataList;
    }

    /**
     * 创建偶数行的样式
     * @author 徐明龙 XuMingLong 2023-03-07
     * @param rowNo
     * @param maxColumnSize
     * @return java.util.List<com.feiyizhan.excel.export.utils.data.ExcelCellCustomData>
     */
    private List<ExcelCellCustomData> createEvenRowStyle(int rowNo,int maxColumnSize){
        ExcelCellFormatter customCellFormatter =(cell)->{
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cell.getSheet().getWorkbook().createFont();
            font.setBold(true);//粗体显示
            font.setColor(IndexedColors.RED.getIndex());
            font.setFontName("微软雅黑");
            font.setFontHeightInPoints((short) 12); //字体大小
            cellStyle.setFont(font);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置实心填充
            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index); //设置填充的背景色
        };

        ExcelCellFormatter customCellFormatter2 =(cell)->{
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cell.getSheet().getWorkbook().createFont();
            font.setBold(false);//粗体显示
            font.setColor(IndexedColors.YELLOW.getIndex());
            font.setFontName("微软雅黑");
            font.setFontHeightInPoints((short) 12);
            cellStyle.setFont(font);
        };

        ExcelCellFormatter customCellFormatter3 =(cell)->{
            XSSFCellStyle cellStyle = cell.getCellStyle();
            XSSFFont font = cell.getSheet().getWorkbook().createFont();
            font.setBold(false);//粗体显示
            font.setColor(IndexedColors.BLUE.getIndex());
            font.setFontName("微软雅黑");
            font.setFontHeightInPoints((short) 12);
            cellStyle.setFont(font);
        };

        List<ExcelCellCustomData> cellCustomDataList = new ArrayList<>(maxColumnSize);
        int i = 0;
        //序号列的样式(参考行的样式）
        cellCustomDataList.add(ExcelReferCellStyleData.builder()
            .referStyleCell(ExcelCellLocation.builder().row(rowNo).column(i).build())
            .build());
        i++;
        //客户编号的样式 (参考行的样式，但创建新不共享的样式)
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .referStyleCell(ExcelCellLocation.builder().row(rowNo).column(i).build())
            .customFormatterList(Arrays.asList(customCellFormatter))
            .build());
        i++;
        //客户名称的样式 （不参考行的样式，创建不共享的样式）
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .build());
        i++;
        //统一社会信用代码的样式 （参考行的样式，并设置值的类型为单行文本）
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .customFormatterList(Arrays.asList(customCellFormatter3))
            .customCellValueFun(SINGLE_LINE_TEXT.getSetValueFun())
            .build());
        i++;
        //客户的备注的样式 （参考行的样式，并设置值的类型为多行文本）
        cellCustomDataList.add(ExcelCustomValueAndStyleData.builder()
            .customFormatterList(Arrays.asList(customCellFormatter2,MULTI_LINE_TEXT.getCellFormatter()))
            .customCellValueFun(MULTI_LINE_TEXT.getSetValueFun())
            .build());
        i++;
        for(;i<maxColumnSize;i++){
            //之后的所有列的样式 （参考行的样式)
            cellCustomDataList.add(ExcelReferCellStyleData.builder()
                .referStyleCell(ExcelCellLocation.builder().row(rowNo).column(i).build())
                .build());
        }
        return cellCustomDataList;
    }
    /**
     * 填充模版变量Map
     * @author 徐明龙 XuMingLong 2023-03-06
     * @param variableMap
     * @return void
     */
    private void fillVariableMap(Map<String, ExcelCellData> variableMap){
        variableMap.put("date",ExcelCellData.builder().value("2023-03-06").build());
        variableMap.put("time",ExcelCellData.builder().value("14:45:40").build());
        variableMap.put("sign",ExcelCellData.builder().value("徐明龙").build());
        //自定义固定值的样式
        variableMap.put("test2",ExcelCellData.builder().value("测试2").cellCustomData(
            ExcelCustomValueAndStyleData.builder().customFormatterList(
                Arrays.asList(
                    (cell)->{
                        XSSFCellStyle cellStyle = cell.getCellStyle();
                        XSSFFont font = cell.getSheet().getWorkbook().createFont();
                        font.setBold(false);//粗体显示
                        font.setColor(IndexedColors.RED.getIndex());
                        cellStyle.setFont(font);
                    }
                )).build()
        ).build());
        variableMap.put("test3",ExcelCellData.builder().value("测试3").build());
    }

//    /**
//     * 测试自动创建sheet并替换导出
//     * @author 徐明龙 XuMingLong 2021-01-28
//     */
//    @Test
//    public void test_replaceAndExportExcelWithAutoCreateSheet(){
//        ExcelData excelData  = new ExcelData();
//        excelData.setTemplateSheetIndex(0);
//        List<ExcelTemplateData> sheetDataList = new ArrayList<>();
//
//        ExcelTemplateData templateData = new ExcelTemplateData();
//        List<List<ExcelCellData>> dataList = new ArrayList<>();
//        Map<String, ExcelCellData> fixDataMap = new HashMap<>();
//        templateData.setBeginRow(1);
//        templateData.setDataList(dataList);
//        templateData.setFixDataMap(fixDataMap);
//        templateData.setSheetName("模版1");
//
//
//        BiConsumer<XSSFCellStyle, XSSFWorkbook> customCellStyleFun1 = (wrapTextStyle,xSSFWorkbook)->{
//            wrapTextStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//            Font font =xSSFWorkbook.createFont();
//            font.setColor(IndexedColors.RED.index);
//            wrapTextStyle.setFont(font);
//            wrapTextStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
//
//        };
//
//        BiConsumer<XSSFCellStyle, XSSFWorkbook> customCellStyleFun2 = (wrapTextStyle,xSSFWorkbook)->{
//            wrapTextStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//            wrapTextStyle.setFillForegroundColor(IndexedColors.BLUE.index);
//        };
////        BiConsumer<XSSFCellStyle, XSSFWorkbook> customCellStyleFun2 = null;
//        for(int i =0 ;i<2000;i++){
//            List<ExcelCellData> baseRowList1 = new ArrayList<>();
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_NUMBER, dataList.size() + 1));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "aaaa "
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList1.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            dataList.add(baseRowList1);
//
//            List<ExcelCellData> baseRowList2 = new ArrayList<>();
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_NUMBER, dataList.size() + 1));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT,
//                "我是谁ssssssssssss", customCellStyleFun1));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12", customCellStyleFun2));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT,
//                "我是谁ssssssssssss", customCellStyleFun1));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12", customCellStyleFun2));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT,
//                "我是谁ssssssssssss", customCellStyleFun1));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12", customCellStyleFun2));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT,
//                "我是谁ssssssssssss", customCellStyleFun1));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12", customCellStyleFun2));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT,
//                "我是谁ssssssssssss", customCellStyleFun1));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12", customCellStyleFun2));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT,
//                "我是谁ssssssssssss", customCellStyleFun1,ExcelCellData.CREATE_CLONE_NEW_CELL_STYLE_FUN));
//            baseRowList2.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12", customCellStyleFun2,
//                ExcelCellData.CREATE_CLONE_NEW_CELL_STYLE_FUN));
//            dataList.add(baseRowList2);
//
//            List<ExcelCellData> baseRowList3 = new ArrayList<>();
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_NUMBER, dataList.size() + 1));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "sdsad "
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "sdsad "
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "sdsad "
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "sdsad "
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "sdsad "
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "sdsad "
//            ));
//            baseRowList3.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"
//            ));
//            dataList.add(baseRowList3);
//        }
//
//
//
//        sheetDataList.add(templateData);
//        excelData.setSheetDataList(sheetDataList);
//        System.out.println(sheetDataList.size());
//        String outFile = "D:\\test.xlsx";
//        //        String template = "d:\\template.xlsx";
//        //        String template = "d:\\bill_statement_no_invoice_list.xlsx";
//        String template = "D:\\bill_statement_customer_sale_list.xlsx";
//        try {
//            FileOutputStream out = new FileOutputStream(outFile);
//            ExportExcelUtil.replaceAndExportExcel(excelData
//                , new FileInputStream(template),out);
//        } catch (FileNotFoundException ex) {
//            ex.printStackTrace();
//        }
//    }
//
//
//    /**
//     * 测试自动创建sheet并替换导出
//     * @author 徐明龙 XuMingLong 2021-01-28
//     */
//    @Test
//    public void test_replaceAndExportExcelWithAutoCreateSheet2(){
//        String template = "d:\\bill_statiment_customer_sale_list.xlsx";
//        String outFile = "d:\\test.xlsx";
//        ExcelData excelData  = new ExcelData();
//        excelData.setTemplateFileName(template);
//        List<ExcelTemplateData> sheetDataList = new ArrayList<>();
//        excelData.setSheetDataList(sheetDataList);
//        excelData.setTemplateSheetIndex(0);
//        excelData.setOutFileName(outFile);
//        for(int i=0;i<3;i++){
//            ExcelTemplateData templateData = new ExcelTemplateData();
//            List<List<ExcelCellData>> dataList = new ArrayList<>();
//            Map<String, ExcelCellData> fixDataMap = new HashMap<>();
//            templateData.setBeginRow(1);
//            templateData.setDataList(dataList);
//            templateData.setFixDataMap(fixDataMap);
//            templateData.setSheetName("模版"+i);
//            List<ExcelCellData> baseRowList = new ArrayList<>();
//            baseRowList.add(new ExcelCellData(ExcelCellData.TYPE_NUMBER, dataList.size() + 1));
//            baseRowList.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "猜猜我是谁"));
//            baseRowList.add(new ExcelCellData(ExcelCellData.TYPE_SINGLE_LINE_TEXT, "2018-08-12"));
//            dataList.add(baseRowList);
//            sheetDataList.add(templateData);
//        }
//
//        try {
//            FileOutputStream out = new FileOutputStream(outFile);
//            ExportExcelUtil.replaceAndExportExcelWithAutoCreateSheet(excelData
//                , new FileInputStream(template),out);
//        } catch (FileNotFoundException ex) {
//            ex.printStackTrace();
//        }
//    }
}
