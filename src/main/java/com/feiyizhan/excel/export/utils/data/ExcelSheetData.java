/**
 *
 */
package com.feiyizhan.excel.export.utils.data;

import lombok.Data;

import java.util.List;

/**
 * Excel Sheet 的数据
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
public class ExcelSheetData {

    /**
     * 使用模版的Sheet索引(从0开始）
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private int templateSheetIndex;


    /**
     * 输出的Sheet名称
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private String outSheetName;

    /**
     * 表格列表，多个二维表
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private List<ExcelTableData> tableDataList ;

    /**
     * 是否为空表格
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private boolean emptied;


}
