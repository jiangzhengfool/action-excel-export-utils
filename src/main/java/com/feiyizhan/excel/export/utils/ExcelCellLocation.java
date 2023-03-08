package com.feiyizhan.excel.export.utils;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Excel 单元格位置
 * @author 徐明龙 XuMingLong 2023-03-02
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelCellLocation {
    /**
     * 行号 （起始下标为0）
     * @author 徐明龙 XuMingLong 2023-03-02
     */
    private int row;
    /**
     * 列号 （起始下标为0）
     * @author 徐明龙 XuMingLong 2023-03-02
     */
    private int column;

}
