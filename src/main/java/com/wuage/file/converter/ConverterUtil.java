package com.wuage.file.converter;

import java.util.HashSet;
import java.util.Set;

public class ConverterUtil {

    private final static Set<String> NEED_CONVERTER_FILE_EXTENSION_SET = new HashSet<String>() {

        private static final long serialVersionUID = 1L;

        {
            this.add("xls");
            this.add("csv");
            this.add("xlsx");
            this.add("xlsm");
            this.add("doc");
            this.add("docx");
            this.add("txt");
            this.add("ppt");
            this.add("pptx");
        }
    };

    /**
     * 判断是否需要转换
     * 
     * @param extension
     * @return
     */
    public static boolean needConverter(String extension) {
        if (extension == null || "".equals(extension)) {
            return false;
        }
        return NEED_CONVERTER_FILE_EXTENSION_SET.contains(extension);
    }
}
