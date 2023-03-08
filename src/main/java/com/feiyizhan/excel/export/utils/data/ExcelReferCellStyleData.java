package com.feiyizhan.excel.export.utils.data;

import com.feiyizhan.excel.export.utils.ExcelCellFormatter;
import com.feiyizhan.excel.export.utils.ExcelCellLocation;
import com.feiyizhan.excel.export.utils.ExcelUtil;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.function.BiConsumer;

/**
 * Excel 参考单元格样式数据
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelReferCellStyleData implements ExcelCellCustomData {

    /**
     * 自定义单元格的值的方法
     * @author 徐明龙 XuMingLong 2023-03-07
     */
    private BiConsumer<XSSFCell, Object> customCellValueFun;

    /**
     * 参考样式的单元格
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private ExcelCellLocation referStyleCell;

    /**
     * 获取单元格的格式化处理器
     * @author 徐明龙 XuMingLong 2023-03-03
     * @return org.apache.poi.xssf.usermodel.XSSFCellStyle
     */
    @Override public ExcelCellFormatter getCellFormatter() {
        return (cell)->{
            XSSFSheet sheet = cell.getSheet();
            XSSFCell referCell = sheet.getRow(referStyleCell.getRow()).getCell(referStyleCell.getColumn());
            if(referCell==null){
                referCell = sheet.getRow(referStyleCell.getRow()).createCell(referStyleCell.getColumn());
            }
            cell.setCellStyle(ExcelUtil.createShareCellStyle(referCell));
        };
    }


    /**
     * 获取设置单元格值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     * @return java.util.function.BiConsumer<org.apache.poi.xssf.usermodel.XSSFCell, java.lang.Object>
     */
    @Override public BiConsumer<XSSFCell, Object> getSetValueFun() {
        return this.customCellValueFun==null?DEFAULT_SET_VALUE_FUN:this.customCellValueFun;
    }
}
