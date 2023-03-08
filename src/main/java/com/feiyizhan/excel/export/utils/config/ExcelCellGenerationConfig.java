package com.feiyizhan.excel.export.utils.config;

import com.feiyizhan.excel.export.utils.ExcelCellFormatter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.function.BiConsumer;

/**
 * Excel单元格生成配置
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class ExcelCellGenerationConfig {
    /**
     * 单元格的值
     * @author 徐明龙 XuMingLong
     * @createDate 2023-03-01
     */
    private Object value;

    /**
     * 设置单元格值的方法
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private BiConsumer<XSSFCell,Object> setValueFun;

    /**
     * 格式化处理器
     * @author 徐明龙 XuMingLong 2023-03-03
     */
    private ExcelCellFormatter formatter;


}
