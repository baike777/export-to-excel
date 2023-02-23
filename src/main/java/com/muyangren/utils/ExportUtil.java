package com.muyangren.utils;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import com.muyangren.excel.ExcelExportStyler;
import org.apache.poi.ss.usermodel.*;


import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;
import java.util.List;

/**
 * @author: muyangren
 * @Date: 2023/2/14
 * @Description: 动态导出工具
 * @Version: 1.0
 */
public class ExportUtil {

    /**
     * 通用
     *
     * @param title      下载文件名
     * @param entityList 动态列
     * @param listMap    数据
     * @param response   通过浏览器下载
     * @param isBrowser  是否通过浏览器下载 true-是 false-否
     */
    public static void dynamicExport(String title, List<ExcelExportEntity> entityList, ExportParams exportParams, List<HashMap<String, Object>> listMap, HttpServletResponse response, boolean isBrowser) {
        exportParams.setStyle(ExcelExportStyler.class);
        //1、默认HSSF的话 创建的就是HSSFWorkbook  wps打开正常，office打不开
        exportParams.setType(ExcelType.XSSF);
        //(选其一)控制标题、表头高度、记得去【ExcelExportStyler】修改下字体大小
        //exportParams.setTitleHeight((short) 12);
        //2、增加数据集合高度(默认即可)
        //exportParams.setHeight((short) 20);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, entityList, listMap);

        //(选其一)实现标题与字段高度不一致、记得去【ExcelExportStyler】修改下字体大小
        //高度自定义样式：
        customStyle(workbook);

        //3、通过浏览器下载(最好是xlsx 不建议xls 特别是做国产化适配 例如麒麟系统 不兼容 xls)
        if (isBrowser) {
            FileUtils.browserDownload(response, title + ".xlsx", workbook);
        } else {
            //下载到本地
            FileUtils.localDownload(title + ".xlsx", workbook);
        }
    }


     private static void customStyle(Workbook workbook){
         assert workbook != null;
         Sheet sheetOne = workbook.getSheetAt(0);
         for (int i = 0; i <= sheetOne.getLastRowNum(); i++) {
             Row row = sheetOne.getRow(i);
             if (i == 0) {
                 //设置第一行的行高（标题）
                 row.setHeightInPoints(55);
             } else if (i == 1) {
                 //设置第二行的行高（表头）
                 row.setHeightInPoints(40);
             } else {
                 int lastCellNum = row.getLastCellNum();
                 for (int j = 0; j < lastCellNum; j++) {
                     //这里我们来实现下奇数填充红色、双数填充蓝色
                     Cell cell = row.getCell(j);
                     if (j % 2 != 0) {
                         //奇数
                         CellStyle titleStyle = workbook.createCellStyle();
                         titleStyle.setFillForegroundColor(IndexedColors.RED1.getIndex());
                         titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                         cell.setCellStyle(titleStyle);
                     } else {
                         //奇数
                         CellStyle titleStyle = workbook.createCellStyle();
                         titleStyle.setFillForegroundColor(IndexedColors.BLUE1.getIndex());
                         titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                         //加粗
                         Font font = workbook.createFont();
                         //字体样式
                         font.setFontName("宋体");
                         //是否加粗
                         font.setBold(true);
                         titleStyle.setFont(font);
                         cell.setCellStyle(titleStyle);
                     }
                 }

             }
         }

     }

}
