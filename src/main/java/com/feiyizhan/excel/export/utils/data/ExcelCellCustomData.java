package com.feiyizhan.excel.export.utils.data;

import com.feiyizhan.excel.export.utils.ExcelCellFormatter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.List;
import java.util.function.BiConsumer;

/**
 * Excel 单元格自定义数据
 * @author 徐明龙 XuMingLong 2023-03-03
 */
public interface ExcelCellCustomData {
    /**
     * 默认的设置值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    BiConsumer<XSSFCell,Object> DEFAULT_SET_VALUE_FUN = (cell,data)->{
        if(data!=null){
            if(data instanceof  Boolean){
                cell.setCellValue(((Boolean)data).booleanValue());
            }else if(data instanceof Number) {
                cell.setCellValue(((Number)data).doubleValue());
            }else if(data instanceof List){
                cell.setCellValue(StringUtils.join((List<String>)data, "\n"));
            }else{
                cell.setCellValue(String.valueOf(data));
            }
        }
    };

    /**
     * 获取设置单元格值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     * @return java.util.function.BiConsumer<org.apache.poi.xssf.usermodel.XSSFCell,java.lang.Object>
     */
    default BiConsumer<XSSFCell,Object> getSetValueFun(){
        return DEFAULT_SET_VALUE_FUN;
    }

    /**
     * 获取单元格的格式化处理器
     * @author 徐明龙 XuMingLong 2023-03-03
     * @return org.apache.poi.xssf.usermodel.XSSFCellStyle
     */
    default ExcelCellFormatter getCellFormatter(){
        return ExcelCellFormatter.DEFAULT_FORMATTER;
    };
}
