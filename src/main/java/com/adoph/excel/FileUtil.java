package com.adoph.excel;

import java.io.InputStream;

/**
 * 工具
 */
public class FileUtil {

    public static InputStream getResourcesFileInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
    }
}
