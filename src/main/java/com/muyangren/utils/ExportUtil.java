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
     * 优质资源审核/推送记录 通用
     *
     * @param title
     * @param entityList
     * @param listMap
     * @param response
     * @param isBrowser 是否通过浏览器下载 1-是 2-否
     */
    public static void dynamicExport(String title, List<ExcelExportEntity> entityList, ExportParams exportParams, List<HashMap<String, Object>> listMap, HttpServletResponse response, boolean isBrowser) {
        exportParams.setStyle(ExcelExportStyler.class);
        //默认HSSF的话 使用office打不开
        exportParams.setType(ExcelType.XSSF);
        //标题高度
        exportParams.setTitleHeight((short) 12);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, entityList, listMap);
        //通过浏览器下载
        if (isBrowser){
            FileUtils.browserDownload(response, title + ".xlsx", workbook);
        }else {
            //下载到本地
            FileUtils.localDownload(title+".xlsx", workbook);
        }
    }

}
