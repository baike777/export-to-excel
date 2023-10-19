package com.muyangren;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author: muyangren
 * @Date: 2023/1/31
 * @Description: easypoi实现多sheet，动态列导出（非注解形式）
 * @Version: 1.0
 */
public class ExportToExcelTest {

    public static void main(String[] args) {
        // 1、创建一个Workbook（XSSFWorkbook）
        Workbook workbook = new XSSFWorkbook();
        // 2、导出sheet1
        exportSheetOne(workbook);
        // 3、导出sheet2
        exportSheetTwo(workbook);
        // 4、通过浏览器或者下载到临时路径
        exportWorkbook(workbook);
    }

    private static void exportSheetOne(Workbook workbook) {
        // 1、创建动态表头
        List<ExcelExportEntity> entityList = new ArrayList<>();
        // 1.1、动态表头
        String sortKey = "colNum";
        for (int i = 0; i < 5; i++) {
            entityList.add(new ExcelExportEntity("列头" + i, sortKey + i, 10));
        }
        // 2、填充数据集合
        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int j = 0; j < 100; j++) {
            HashMap<String, Object> contentMap = new HashMap<>();
            for (int x = 0; x < 5; x++) {
                contentMap.put(sortKey + x, "数据" + x);
            }
            dataList.add(contentMap);
        }
        // 3、定义标题和sheetName
        ExportParams exportParams = new ExportParams("【大标题】", "sheet1");
        // 默认未 ExcelType.HSSF导出失败
        exportParams.setType(ExcelType.XSSF);
        ExcelExportService service = new ExcelExportService();
        service.createSheetForMap(workbook, exportParams, entityList, dataList);
    }

    private static void exportSheetTwo(Workbook workbook) {
        // 1、创建动态表头
        List<ExcelExportEntity> entityList = new ArrayList<>();
        // 1.1、动态表头
        String sortKey = "colNum";
        for (int i = 0; i < 3; i++) {
            entityList.add(new ExcelExportEntity("列头" + i, sortKey + i, 10));
        }
        // 2、填充数据集合
        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int j = 0; j < 3; j++) {
            HashMap<String, Object> contentMap = new HashMap<>();
            for (int x = 0; x < 3; x++) {
                contentMap.put(sortKey + x, "sheet2数据" + x);
            }
            dataList.add(contentMap);
        }
        // 3、定义标题和sheetName
        ExportParams exportParams = new ExportParams(null, "sheet2");
        // 默认未 ExcelType.HSSF导出失败
        exportParams.setType(ExcelType.XSSF);
        ExcelExportService service = new ExcelExportService();
        service.createSheetForMap(workbook, exportParams, entityList, dataList);
    }

    private static void exportWorkbook(Workbook workbook) {
        //1、通过浏览器下载(一般实际项目中用到)
        //ServletOutputStream outputStream = response.getOutputStream();
        //response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
        //response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode("动态列多sheet.xls", "UTF-8"));
        //workbook.write(outputStream);
        //outputStream.flush();
        //outputStream.close();

        // 2、直接下载到本地(一般用于测试)
        //注*：建议最好就是【.xlsx】后缀
        try (FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\WUDI\\Desktop\\导出测试用例.xlsx");) {
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
