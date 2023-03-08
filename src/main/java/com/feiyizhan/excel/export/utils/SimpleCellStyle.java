package com.feiyizhan.excel.export.utils;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 单元格样式
 * @author 邱光益 qiugy 2023/1/9
 */
@Data
@AllArgsConstructor
@Builder
public class SimpleCellStyle {

    /**
     * 字体加粗
     * @author 邱光益 qiugy 2023-01-09
     **/
    private Boolean fontBold;
    /**
     * 字体颜色
     * @author 邱光益 qiugy 2023-01-09
     **/
    private IndexedColors fontColor;
    /**
     * 背景色
     * @author 邱光益 qiugy 2023-01-09
     **/
    private IndexedColors groundColor;
}
