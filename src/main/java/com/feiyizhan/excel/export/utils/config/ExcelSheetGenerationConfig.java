package com.feiyizhan.excel.export.utils.config;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

/**
 * Excel 页生成配置
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelSheetGenerationConfig {

    /**
     * 新的Sheet
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private XSSFSheet newSheet;

    /**
     * 输出的Sheet名称
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private String outSheetName;

    /**
     * 表格列表
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private List<ExcelTableGenerationConfig> tableList;

}
