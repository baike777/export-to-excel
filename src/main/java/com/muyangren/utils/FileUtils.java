package com.muyangren.utils;

import org.apache.poi.ss.usermodel.Workbook;

import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.rmi.ServerException;
import java.util.List;
import java.util.Objects;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 处理文件工具
 */
public class FileUtils {



    /**
     * 下载到本地路径
     * @param file
     * @return
     * @throws IOException
     */
    public static File fileDownloadToLocalPath(MultipartFile file) {
        File destFile = null;
        try {
            //获取文件名称
            if (file.getOriginalFilename() == null || Objects.equals(file.getOriginalFilename(), "")){
                throw new ServerException("导入案例失败！");
            }
            String fileName = file.getOriginalFilename();
            //获取文件后缀
            String pref = fileName.lastIndexOf(".") != -1 ? fileName.substring(fileName.lastIndexOf(".") + 1) : null;
            //临时文件
            //临时文件名避免重复
            String uuidFile = UUID.randomUUID().toString().replace("-", "") + "." + pref;
            destFile = new File(FileUtils.getProjectPath() + uuidFile);
            if (!destFile.getParentFile().exists()) {
                destFile.getParentFile().mkdirs();
            }
            file.transferTo(destFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return destFile;
    }

    /**
     * 先压缩到临时路径才输出浏览器
     * @param files
     * @param zipName
     * @param filePath
     * @return
     * @throws Exception
     */
    public static String createZipAndReturnPath(List<File> files, String zipName, String filePath)throws Exception{
        if (files.isEmpty()) {
            return null;
        }
        // 创建 FileOutputStream 对象
        FileOutputStream fileOutputStream = null;
        // 创建 ZipOutputStream
        ZipOutputStream zipOutputStream = null;
        // 创建 FileInputStream 对象
        FileInputStream fileInputStream = null;
        // 压缩文件保存路径
        String zipFile = filePath + zipName;
        try {
            // 实例化 FileOutputStream 对象
            fileOutputStream = new FileOutputStream(zipFile);
            // 实例化 ZipOutputStream 对象
            zipOutputStream = new ZipOutputStream(fileOutputStream);
            // 创建 ZipEntry 对象
            ZipEntry zipEntry = null;
            // 遍历源文件数组
            for (File file : files) {
                // 将源文件数组中的当前文件读入 FileInputStream 流中
                fileInputStream = new FileInputStream(file);
                // 实例化 ZipEntry 对象，源文件数组中的当前文件
                zipEntry = new ZipEntry(file.getName());
                zipOutputStream.putNextEntry(zipEntry);
                // 该变量记录每次真正读的字节个数
                int len;
                // 定义每次读取的字节数组
                byte[] buffer = new byte[1024];
                while ((len = fileInputStream.read(buffer)) > 0) {
                    zipOutputStream.write(buffer, 0, len);
                }
                fileInputStream.close();
            }
           return zipFile;
        } catch (IOException e) {
            throw new Exception("压缩案例记录异常：" + e.getMessage());
        }finally {
            assert zipOutputStream != null;
            zipOutputStream.closeEntry();
            zipOutputStream.close();
            assert fileInputStream != null;
            fileInputStream.close();
            fileOutputStream.close();
        }
    }

    public static void downLoadFile(String zipFile, HttpServletResponse response) throws IOException {
        BufferedInputStream bis =null;
        BufferedOutputStream bos =null;
        try {
            File file = new File(zipFile);
            response.reset();
            response.setContentType("application/octet-stream");
            response.setHeader("Content-Disposition", "attachment;filename=\"" + new String(file.getName().getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1) + "\"");
            response.setContentLength((int) file.length());
            bis = new BufferedInputStream(new FileInputStream(file), 4096);
            bos =  new BufferedOutputStream(
                    response.getOutputStream());
            byte[] temp = new byte[4096];
            int i;
            int all = 0;
            while ((i = bis.read(temp)) != -1) {
                bos.write(temp, 0, i);
                all += i;
            }
            bos.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            assert bos != null;
            bos.close();
            assert bis != null;
            bis.close();
        }
    }

    /**
     * 压缩后直接输出到浏览器
     * @param fileList
     * @param zipFileName
     * @param request
     * @param response
     */
    public static void downLoadZip(List<File> fileList, String zipFileName, HttpServletRequest request, HttpServletResponse response) {
        byte[] buf = new byte[1024];
        // 获取输出流
        BufferedOutputStream bos = null;
        try {
            bos = new BufferedOutputStream(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
        FileInputStream in = null;
        ZipOutputStream out = null;
        try {
            // 重置
            response.reset();
            String fileName = "";
            String agent = request.getHeader("user-agent");
            if (agent.contains("FireFox")) {
                fileName = new String(zipFileName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
            } else {
                fileName = URLEncoder.encode(zipFileName, "UTF-8");
            }

            // 不同类型的文件对应不同的MIME类型
            response.setContentType("application/x-msdownload");
            response.setCharacterEncoding("utf-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName );

            // ZipOutputStream类：完成文件或文件夹的压缩
            out = new ZipOutputStream(bos);
            for (int i = 0; i < fileList.size(); i++) {
                in = new FileInputStream(fileList.get(i));
                // 给列表中的文件单独命名
                out.putNextEntry(new ZipEntry(fileList.get(i).getName()));
                int len;
                while ((len = in.read(buf)) != -1) {
                    out.write(buf, 0, len);
                }
                in.close();
            }
            out.closeEntry();
            out.close();
            bos.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (in != null) {
                    in.close();
                }
                if (out != null) {
                    out.close();
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 通过浏览器下载文件
     * @param response
     * @param fileName
     * @param workbook
     */
    public static void browserDownload(HttpServletResponse response, String fileName, Workbook workbook) {
        OutputStream outputStream = null;
        try {
            response.reset();
            response.setContentType("application/octet-stream; charset=utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName,"utf8"));
            outputStream = response.getOutputStream();
            workbook.write(outputStream);
            outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
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
        FileOutputStream outputStream = null;
        try {
            String filePath = getProjectPath() + fileName;
            outputStream = new FileOutputStream(filePath);
            workbook.write(outputStream);
            outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
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
     * @return 文件路径
     */
    public static String getProjectPath(){
        String os = System.getProperty("os.name").toLowerCase();
        //windows下
        if (os.indexOf("windows")>=0) {
            return "C://temp/";
        }else{
            return "/usr/local/temp/";
        }
    }

    /**
     * @return 资源下发模板专用路径
     */
    public static String getResourcesFilePath(){
        return "document" + File.separator + "export" + File.separator + "distribute" + File.separator;
    }
}
