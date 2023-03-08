package com.feiyizhan.excel.export.utils.config;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Excel 行生成配置
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelRowGenerationConfig {

    /**
     * 一行的单元格列表
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private List<ExcelCellGenerationConfig> cellList;
}
