package com.feiyizhan.excel.export.utils.data;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Excel 行数据
 * @author 徐明龙 XuMingLong 2023-03-02
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelRowData {

    /**
     * 单元格数据列表
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private List<ExcelCellData> cellDataList;


}
