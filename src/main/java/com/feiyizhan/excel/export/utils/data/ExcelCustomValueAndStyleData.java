package com.feiyizhan.excel.export.utils.data;

import com.feiyizhan.excel.export.utils.ExcelCellFormatter;
import com.feiyizhan.excel.export.utils.ExcelCellLocation;
import com.feiyizhan.excel.export.utils.ExcelUtil;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;
import java.util.function.BiConsumer;

/**
 * Excel 自定义单元格数据和样式的数据
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelCustomValueAndStyleData implements ExcelCellCustomData {

    /**
     * 自定义单元格的值的方法
     * @author 徐明龙 XuMingLong 2023-03-07
     */
    private BiConsumer<XSSFCell, Object> customCellValueFun;

    /**
     * 自定义单元格样式的方法的列表
     * @author 徐明龙 XuMingLong 2023-03-01
     **/
    private List<ExcelCellFormatter> customFormatterList;

    /**
     * 参考样式的单元格（动态的，以最终的值表格为准）
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private ExcelCellLocation referStyleCell;

    /**
     * 获取单元格样式
     * @author 徐明龙 XuMingLong 2023-03-03
     * @param cell
     * @return org.apache.poi.xssf.usermodel.XSSFCellStyle
     */
     private XSSFCellStyle getCellStyle(XSSFCell cell) {
        XSSFSheet sheet = cell.getSheet();
        XSSFCellStyle cellStyle ;
        if(referStyleCell!=null){
            //根据参考的单元格Style 创建一个不共享的Style
            XSSFCell referCell = sheet.getRow(referStyleCell.getRow()).getCell(referStyleCell.getColumn());
            if(referCell==null){
                referCell = sheet.getRow(referStyleCell.getRow()).createCell(referStyleCell.getColumn());
            }
            cellStyle =
                ExcelUtil.createNotShareCellStyle(referCell);
        }else{
            //创建一个新的不共享的Style
            cellStyle = ExcelUtil.createNotShareCellStyle(sheet.getRow(cell.getRowIndex())
                .getCell(cell.getColumnIndex()));
            }
        return cellStyle;
    }

    /**
     * 获取单元格的格式化处理器
     * @author 徐明龙 XuMingLong 2023-03-03
     * @return org.apache.poi.xssf.usermodel.XSSFCellStyle
     */
    @Override public ExcelCellFormatter getCellFormatter() {
        return (cell)->{
            cell.setCellStyle(getCellStyle(cell));
            //执行自定义样式的处理
            if(CollectionUtils.isNotEmpty(customFormatterList)){
                customFormatterList.forEach(item->{
                    if(item!=null){
                        item.format(cell);
                    }
                });
            }

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
