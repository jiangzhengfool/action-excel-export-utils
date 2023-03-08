package com.feiyizhan.excel.export.utils;

import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.*;
import java.net.URL;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

/**
 * 文件工具类
 * @author 徐明龙 XuMingLong 2023-03-08
 */
public class FileUtil {

    private FileUtil() {
        throw new IllegalStateException("Utility class");
    }

    /**
     * 获取指定文件的输入流
     * @author 徐明龙 XuMingLong 2023-03-08
     * @param fileName
     * @return java.io.InputStream
     */
    public static InputStream getFileInputStream(String fileName){
        Resource resource = new ClassPathResource(fileName);
        File templateFile = getFile(fileName);
        try {
            return templateFile.exists()?new FileInputStream(templateFile):resource.getInputStream();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取指定文件
     * @author 徐明龙 XuMingLong 2023-03-08
     * @param fileName
     * @return java.io.File
     */
    public static File getFile(String fileName){
        File file= new File(fileName);
        if(file.exists()) {
            return file;
        }
        URL url = FileUtil.class.getClassLoader().getResource(fileName);
        if(url!=null){
            try {
                fileName = URLDecoder.decode(url.getFile(), StandardCharsets.UTF_8.name());
                file= new File(fileName);
            } catch (UnsupportedEncodingException e) {}
        }
        return file;
    }
}
