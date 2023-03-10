package com.feiyizhan.excel.export.utils;

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
}
