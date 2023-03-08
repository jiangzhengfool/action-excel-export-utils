package com.feiyizhan.excel.export.utils;

import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * Excel 单元格格式化处理器
 * @author 徐明龙 XuMingLong 2023-03-03
 */
@FunctionalInterface
public interface ExcelCellFormatter {

    /**
     * 默认的格式化方法
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    ExcelCellFormatter DEFAULT_FORMATTER = cell->{};

    /**
     * 格式化单元格
     * @author 徐明龙 XuMingLong 2023-03-03
     * @param cell
     * @return void
     */
    void format(XSSFCell cell);
}
