package com.feiyizhan.excel.export.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Excel测试类
 * @author 徐明龙 XuMingLong 2023-03-10
 */
public class ExcelTests {

    /**
     * 测试简单生成Excel文件的算法
     * @author 徐明龙 XuMingLong 2023-03-10
     * @return void
     */
    @Test
    public void test_AlgorithmForGeneratingSimpleExcel(){
        try {
            //创建一个空白的Excel文件
            XSSFWorkbook workbook = new XSSFWorkbook();
            System.out.println("初始的样式个数："+workbook.getNumCellStyles());
            XSSFSheet sheet = workbook.createSheet("测试Sheet1");
            //填充表格
            for(int rowNo=0;rowNo<10;rowNo++){
                XSSFRow row = sheet.createRow(rowNo);
                //创建单元格
                for(int cellNo=0;cellNo<10;cellNo++){
                    XSSFCell cell = row.createCell(cellNo);
                    //设置每个单元格的值，并设置每个单元格的样式
                    XSSFCellStyle cellStyle = workbook.createCellStyle();
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置实心填充
                    cellStyle.setFillForegroundColor(IndexedColors.BLUE1.index); //设置填充的背景色
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(((char)(rowNo+'A'))+ ":"+ (cellNo+1));
                }
            }
            System.out.println("最终的样式个数："+workbook.getNumCellStyles());
            String outFile = "D:\\test_AlgorithmForGeneratingSimpleExcel_1.xlsx";
            workbook.write(new FileOutputStream(outFile));
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 测试简单生成Excel文件的算法2
     * @author 徐明龙 XuMingLong 2023-03-10
     * @return void
     */
    @Test
    public void test_AlgorithmForGeneratingSimpleExcel2(){
        try {
            //创建一个空白的Excel文件
            XSSFWorkbook workbook = new XSSFWorkbook();
            System.out.println("初始的样式个数："+workbook.getNumCellStyles());
            XSSFSheet sheet = workbook.createSheet("测试Sheet1");
            //提前创建好单元格样式
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置实心填充
            cellStyle.setFillForegroundColor(IndexedColors.BLUE1.index); //设置填充的背景色

            //填充表格
            for(int rowNo=0;rowNo<10;rowNo++){
                XSSFRow row = sheet.createRow(rowNo);
                //创建单元格
                for(int cellNo=0;cellNo<10;cellNo++){
                    XSSFCell cell = row.createCell(cellNo);
                    //设置每个单元格的值，并设置每个单元格的样式
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(((char)(rowNo+'A'))+ ":"+ (cellNo+1));
                }
            }
            System.out.println("最终的样式个数："+workbook.getNumCellStyles());
            String outFile = "D:\\test_AlgorithmForGeneratingSimpleExcel_2.xlsx";
            workbook.write(new FileOutputStream(outFile));
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 测试Shift Row
     * @author 徐明龙 XuMingLong 2023-03-10
     * @return void
     */
    @Test
    public void test_shiftRows1(){
        try {
            String templateFile = "template_other.xlsx";
            String outFile = "D:\\test_shiftRows1.xlsx";
            XSSFWorkbook wb = new XSSFWorkbook(FileUtil.getFile(templateFile));
            XSSFSheet newSheet = wb.cloneSheet(0);
            int lastRowNum = newSheet.getLastRowNum();
            newSheet.shiftRows(2,lastRowNum,10);
            wb.removeSheetAt(0);
            wb.write(new FileOutputStream(outFile));
            /**
             * 存在的BUG，
             * 1、shiftRows 对于4.1.1版本之后的POI会导致Excel文件打开报错。
             * 2、shiftRows 执行后，再执行XSSFWorkbook.write输出到另外一个文件，会导致原文件也被修改了行号。
             * 下次再打开源文件会出现行号是移动后的行号，但用Excel让软件打开则是正常的没有移动行号。
             */
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

    }


    /**
     * 测试Shift Row
     * @author 徐明龙 XuMingLong 2023-03-10
     * @return void
     */
    @Test
    public void test_shiftRows2(){
        try {
            String templateFile = "template_other.xlsx";
            String outFile = "D:\\test_shiftRows2.xlsx";
            XSSFWorkbook wb = new XSSFWorkbook(FileUtil.getFile(templateFile));
            XSSFSheet templateSheet =  wb.getSheetAt(0);
            System.out.println("shiftRows前模版的最大行号:"+templateSheet.getLastRowNum());
            templateSheet.rowIterator().forEachRemaining(item->{
                System.out.println("当前行号："+item.getRowNum());
            });
            XSSFSheet newSheet = wb.cloneSheet(0);
            int lastRowNum = newSheet.getLastRowNum();
            newSheet.shiftRows(2,lastRowNum,10);
            wb.removeSheetAt(0);
            wb.write(new FileOutputStream(outFile));
            System.out.println("shiftRows后模版的最大行号:"+templateSheet.getLastRowNum());
            templateSheet.rowIterator().forEachRemaining(item->{
                System.out.println("当前行号："+item.getRowNum());
            });
            wb.close();
            /**
             * 存在的BUG，
             * 1、shiftRows 对于4.1.1版本之后的POI会导致Excel文件打开报错。
             * 2、shiftRows 执行后，再执行XSSFWorkbook.write输出到另外一个文件，会导致原文件也被修改了行号。
             * 下次再打开源文件会出现行号是移动后的行号，但用Excel让软件打开则是正常的没有移动行号。
             * 通过Excel软件打开模版文件后就，选择行号异常的行区域，执行取消隐藏操作后，就可以恢复正常行号。
             */
            wb = new XSSFWorkbook(FileUtil.getFile(templateFile));
            templateSheet =  wb.getSheetAt(0);
            System.out.println("重新打开模版的最大行号:"+templateSheet.getLastRowNum());
            templateSheet.rowIterator().forEachRemaining(item->{
                System.out.println("当前行号："+item.getRowNum());
            });
            wb.close();


        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

    }
}
