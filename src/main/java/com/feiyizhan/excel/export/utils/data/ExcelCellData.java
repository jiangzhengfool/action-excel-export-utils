package com.feiyizhan.excel.export.utils.data;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Excel 单元格数据
 * @author 徐明龙 XuMingLong 2023-03-02
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelCellData {

    /**
     * 单元格的值
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private Object value;

    /**
     * 单元格自定义数据
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private ExcelCellCustomData celCustomData;



}
