package com.feiyizhan.excel.export.utils.data;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.List;
import java.util.Map;

/**
 * Excel文件数据对象
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Setter@Getter@ToString @EqualsAndHashCode
public class ExcelFileData {
    /**
     * 输出的文件名
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private String outFileName;


    /**
     * 每个Sheet的数据
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private List<ExcelSheetData> sheetDataList;

    /**
     * 变量数据
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private Map<String, ExcelCellData> variableMap;

}

