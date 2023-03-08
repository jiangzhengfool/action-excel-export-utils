package com.feiyizhan.excel.export.utils.config;

import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * Excel 文件生成配置
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
public class ExcelWorkBookGenerationConfig {

    /**
     * Excel 文件
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private XSSFWorkbook workbook ;

    /**
     * 模版的Sheet列表
     * @author 徐明龙 XuMingLong 2023-03-01
     **/
    private List<XSSFSheet> templateSheetList;

    /**
     * 输出的文件名
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private String outFileName;

    /**
     * Sheet列表
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private List<ExcelSheetGenerationConfig> sheetList;

    /**
     * 变量数据
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private Map<String, ExcelCellGenerationConfig> variableMap;
}
