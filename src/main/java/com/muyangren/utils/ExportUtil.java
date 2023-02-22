package com.muyangren.utils;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import com.muyangren.excel.ExcelExportStyler;
import org.apache.poi.ss.usermodel.Workbook;


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
     * @param title  下载文件名
     * @param entityList 动态列
     * @param listMap  数据
     * @param response 通过浏览器下载
     * @param isBrowser 是否通过浏览器下载 true-是 false-否
     */
    public static void dynamicExport(String title, List<ExcelExportEntity> entityList, ExportParams exportParams, List<HashMap<String, Object>> listMap, HttpServletResponse response, boolean isBrowser) {
        exportParams.setStyle(ExcelExportStyler.class);
        //默认HSSF的话 创建的就是HSSFWorkbook  wps打开正常，office打不开
        exportParams.setType(ExcelType.XSSF);
        //大标题小标题高度
        exportParams.setTitleHeight((short) 12);
        //增加数据集合高度(默认即可)
        //exportParams.setHeight((short) 20);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, entityList, listMap);
        //通过浏览器下载(最好是xlsx 不建议xls 特别是做国产化适配 例如麒麟系统 不兼容 xls)
        if (isBrowser){
            FileUtils.browserDownload(response, title + ".xlsx", workbook);
        }else {
            //下载到本地
            FileUtils.localDownload(title+".xlsx", workbook);
        }
    }

}
