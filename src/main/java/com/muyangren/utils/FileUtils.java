package com.muyangren.utils;

import org.apache.poi.ss.usermodel.Workbook;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
/**
 * 处理文件工具
 * @author guangsheng
 */
public class FileUtils {

    private FileUtils() {
    }

    /**
     * 定义文件路径属性
     */
    public static final String WINDOWS_PATH = "C:\\Users\\WUDI\\Desktop\\";
    public static final String LINUX_PATH = "/usr/local/temp/";

    /**
     * 通过浏览器下载文件
     * @param response
     * @param fileName
     * @param workbook
     */
    public static void browserDownload(HttpServletResponse response, String fileName, Workbook workbook) {
        try (OutputStream outputStream = response.getOutputStream()) {
            response.reset();
            response.setContentType("application/octet-stream; charset=utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "utf-8"));
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 下载到本地路径（一般用于测试）
     * @param fileName
     * @param workbook
     */
    public static void localDownload(String fileName, Workbook workbook) {
        String filePath = getProjectPath() + fileName;
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 删除文件夹/文件
     *
     * @param directory 要被删除的文件夹
     */
    public static void delAllFile(File directory) {
        if (!directory.isDirectory()) {
            directory.delete();
        } else {
            File[] files = directory.listFiles();
            // 空文件夹
            assert files != null;
            if (files.length == 0) {
                directory.delete();
                return;
            }
            // 删除子文件夹和子文件
            for (File file : files) {
                if (file.isDirectory()) {
                    delAllFile(file);
                } else {
                     file.delete();
                }
            }
            // 删除文件夹自己
            directory.delete();
        }
    }

    /**
     * 获取文件路径的方法
     * @return
     */
    public static String getProjectPath() {
        String os = System.getProperty("os.name").toLowerCase();
        if (os.contains("windows")) {
            return WINDOWS_PATH;
        } else {
            return LINUX_PATH;
        }
    }

}
