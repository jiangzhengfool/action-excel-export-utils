package com.feiyizhan.excel.export.utils.data;

import com.feiyizhan.excel.export.utils.ExcelCellFormatter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.List;
import java.util.function.BiConsumer;

/**
 * Excel 单元格值类型的数据
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelCellValueTypeData implements ExcelCellCustomData {


    /**
     * 数字的设置值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private static final BiConsumer<XSSFCell,Object> TYPE_NUMBER_SET_VALUE_FUN = (cell,data)->{
        if(data!=null){
            cell.setCellValue(Double.parseDouble(data.toString()));
        }
    };

    /**
     * 字符串的设置值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private static final BiConsumer<XSSFCell,Object> TYPE_SINGLE_LINE_TEXT_SET_VALUE_FUN = (cell,data)->{
        if(data!=null){
            cell.setCellValue(String.valueOf(data.toString()));
        }
    };

    /**
     * 数字的设置值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private static final BiConsumer<XSSFCell,Object> TYPE_MULTI_LINE_TEXT_SET_VALUE_FUN = (cell,data)->{
        if(data!=null){
            List<String> strList = (List<String>)data;
            cell.setCellValue(StringUtils.join(strList, "\n"));
        }
    };

    /**
     * 单行文本
     * @author 徐明龙 XuMingLong
     * @createDate 2023-03-01
     */
    public static final int TYPE_SINGLE_LINE_TEXT = 10;

    /**
     * 多行文本
     * @author 徐明龙 XuMingLong
     * @createDate 2023-03-01
     */
    public static final int TYPE_MULTI_LINE_TEXT =11;

    /**
     * 数字
     * @author 徐明龙 XuMingLong
     * @createDate 2023-03-01
     */
    public static final int TYPE_NUMBER = 20 ;


    /**
     * 单行文本样式
     * @author 徐明龙 XuMingLong 2023-03-06
     */
    public static final ExcelCellValueTypeData SINGLE_LINE_TEXT = new ExcelCellValueTypeData(TYPE_SINGLE_LINE_TEXT);
    /**
     * 多行文本的处理
     * @author 徐明龙 XuMingLong 2023-03-06
     */
    public static final ExcelCellValueTypeData MULTI_LINE_TEXT = new ExcelCellValueTypeData(TYPE_MULTI_LINE_TEXT);
    /**
     * 数字的处理
     * @author 徐明龙 XuMingLong 2023-03-06
     */
    public static final ExcelCellValueTypeData NUMBER = new ExcelCellValueTypeData(TYPE_NUMBER);

    /**
     * 值的类型
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private int valueType;

    /**
     * 获取设置单元格值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     * @return java.util.function.BiConsumer<org.apache.poi.xssf.usermodel.XSSFCell, java.lang.Object>
     */
    @Override public BiConsumer<XSSFCell, Object> getSetValueFun() {
        switch(valueType) {
            case TYPE_NUMBER:
                return TYPE_NUMBER_SET_VALUE_FUN;
            case TYPE_SINGLE_LINE_TEXT:
                return TYPE_SINGLE_LINE_TEXT_SET_VALUE_FUN;
            case TYPE_MULTI_LINE_TEXT:
                return TYPE_MULTI_LINE_TEXT_SET_VALUE_FUN;
            default:
                return DEFAULT_SET_VALUE_FUN;
        }
    }

    /**
     * 获取单元格的格式化处理器
     * @author 徐明龙 XuMingLong 2023-03-03
     * @return org.apache.poi.xssf.usermodel.XSSFCellStyle
     */
    @Override public ExcelCellFormatter getCellFormatter() {
        switch(valueType) {
            case TYPE_MULTI_LINE_TEXT:
                return (cell)->{
                    cell.getCellStyle().setWrapText(true);
                };
            default:
                return ExcelCellFormatter.DEFAULT_FORMATTER;
        }
    }

}
