package com.feiyizhan.excel.export.utils;

import com.feiyizhan.excel.export.utils.config.ExcelWorkBookGenerationConfig;
import com.feiyizhan.excel.export.utils.data.ExcelFileData;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.AbstractMap;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 导出excel工具类
 * @author 徐明龙 XuMingLong 2023-03-01
 */
@Log4j2
public class ExportExcelUtil {
    private ExportExcelUtil() {
        throw new IllegalStateException("Utility class");
    }

    /**
     * 变量匹配模式
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private static final Pattern VARIABLE_PATTERN = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);


    /**
     * 默认的导出Excel响应头的Map
     * @author 徐明龙 XuMingLong 2023-03-01
     */
    private static final Map<String,String> DEFAULT_EXPORT_EXCEL_RESP_HEADER_MAP = Stream.of(
        new AbstractMap.SimpleImmutableEntry<>("Cache-Control", "must-revalidate"),
        new AbstractMap.SimpleImmutableEntry<>("Pragma", "public"),
        // 告诉浏览器用什么软件可以打开此文件
        new AbstractMap.SimpleImmutableEntry<>("content-Type", "application/vnd.ms-excel"),
        new AbstractMap.SimpleImmutableEntry<>("Access-Control-Expose-Headers", "Content-Disposition")
        )
        .collect(Collectors.toMap(item->item.getKey(),item->item.getValue()));

    /**
     * 获取指定输出文件名的导出Excel的文件的响应头
     * @author 徐明龙 XuMingLong 2023-03-01
     * @param outFileName
     * @return java.util.Map<java.lang.String,java.lang.String>
     */
    public static Map<String,String> getExportExcelRespHeader(String outFileName){
        Map<String,String> headerMap = new HashMap<>(DEFAULT_EXPORT_EXCEL_RESP_HEADER_MAP);
        // 下载文件的默认名称
        try {
            headerMap.put("Content-Disposition", "attachment;filename=" +
                URLEncoder.encode(outFileName, StandardCharsets.UTF_8.name()).replace("+","%20"));
        } catch (UnsupportedEncodingException e) {
            log.debug("转换输出文件名称异常",e);
            log.debug("使用不转换的文件名称");
            headerMap.put("Content-Disposition", "attachment;filename=" +outFileName);
        }
        return headerMap;
    }

    /**
     * 导出指定Excel到HTTP输出流
     * @author 徐明龙 XuMingLong 2023-03-01
     * @param wb
     * @param response
     * @param outFileName
     * @return void
     */
    public static void exportExcel(XSSFWorkbook wb,HttpServletResponse response,String outFileName) throws IOException {
        //获取导出Excel的响应头
        Map<String,String> headerMap = getExportExcelRespHeader(outFileName);
        //设置响应头
        headerMap.forEach((k,v)->{
            response.setHeader(k,v);
        });
        //输出Excel
        wb.write(response.getOutputStream());
        //刷新输出流
        response.getOutputStream().flush();
    }

    /**
     * 导出Excel文件到输出流
     * @author 徐明龙 XuMingLong 2023-03-06
     * @param wb
     * @param out
     * @return void
     */
    public static void exportExcel(XSSFWorkbook wb,OutputStream out) throws IOException {
        //输出Excel
        wb.write(out);
        //刷新输出流
        out.flush();
    }

    /**
     * 生成并导出Excel文件
     * @author 徐明龙 XuMingLong 2023-03-06
     * @param excelFileData
     * @param templateFileInputStream
     * @param response
     * @return void
     */
    public static void generateAndExportExcel(ExcelFileData excelFileData,
        InputStream templateFileInputStream,HttpServletResponse response){
        ExcelWorkBookGenerationConfig config = ExcelGenerationUtil.generateExcel(excelFileData,templateFileInputStream);
        if(config!=null){
            try {
                exportExcel(config.getWorkbook(),response,config.getOutFileName());
            } catch (IOException e) {
                log.error("导出Excel文件失败",e);
            }
        }
    }

    /**
     * 生成并导出Excel文件到输出流
     * @author 徐明龙 XuMingLong 2023-03-06
     * @param excelFileData
     * @param templateFileInputStream
     * @param out
     * @return void
     */
    public static void generateAndExportExcel(ExcelFileData excelFileData,
        InputStream templateFileInputStream,OutputStream out){
        ExcelWorkBookGenerationConfig config = ExcelGenerationUtil.generateExcel(excelFileData,templateFileInputStream);
        if(config!=null){
            try {
                exportExcel(config.getWorkbook(),out);
            } catch (IOException e) {
                log.error("导出Excel文件失败",e);
            }
        }
    }

    /**
     * 生成并导出Excel到文件
     * @author 徐明龙 XuMingLong 2023-03-06
     * @param excelFileData
     * @param templateFileInputStream
     * @return void
     */
    public static void generateAndExportExcelToFile(ExcelFileData excelFileData,
        InputStream templateFileInputStream){
        ExcelWorkBookGenerationConfig config = ExcelGenerationUtil.generateExcel(excelFileData,templateFileInputStream);
        if(config!=null){
            try {
                exportExcel(config.getWorkbook(),new FileOutputStream(config.getOutFileName()));
            } catch (IOException e) {
                log.error("导出Excel文件失败",e);
            }
        }
    }

    /**
     * 生成并导出Excel到文件
     * @author 徐明龙 XuMingLong 2023-03-06
     * @param excelFileData
     * @param templateFile
     * @return void
     */
    public static void generateAndExportExcelToFile(ExcelFileData excelFileData, File templateFile){
        try {
            ExcelWorkBookGenerationConfig config = ExcelGenerationUtil.generateExcel(excelFileData,
                new FileInputStream(templateFile));
            if(config!=null){
                exportExcel(config.getWorkbook(),new FileOutputStream(config.getOutFileName()));
            }else{
                log.debug("根据Excel模版文件生成新的Excel文件失败处理失败");
            }
        } catch (IOException e) {
            log.error("导出Excel文件失败",e);
        }

    }

}

