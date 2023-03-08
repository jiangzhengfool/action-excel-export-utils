package com.feiyizhan.excel.export.utils.config;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Excel表格生成配置
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelTableGenerationConfig {

    /**
     * 开始行号
     * @author 徐明龙 XuMingLong
     * @createDate 2023-03-01
     */
    private int beginRowNo;

    /**
     * 表格行列表
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private List<ExcelRowGenerationConfig> rowList;
}
