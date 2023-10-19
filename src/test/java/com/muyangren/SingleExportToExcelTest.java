package com.muyangren;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import com.muyangren.excel.utils.ExportUtil;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @author: muyangren
 * @Date: 2023/2/17
 * @Description: 单sheet导出 配合 自定义样式
 * @Version: 1.0
 */
public class SingleExportToExcelTest {
    public static void main(String[] args) {
        //1、填充标题
        String title ="标题";
        //2、填充标题
        List<ExcelExportEntity> resourceEntityList = new ArrayList<>();
        //2.1、动态表头
        String sortKey = "colNum";
        for (int i = 0; i < 5; i++) {
            resourceEntityList.add(new ExcelExportEntity("列头" + i, sortKey + i, 20));
        }
        //3、填充数据集合
        //3.1) 数据数组
        List<HashMap<String, Object>> dataList = new ArrayList<>();
        for (int j = 0; j < 100; j++) {
            //3.2)数据集合
            HashMap<String, Object> contentMap = new HashMap<>();
            //3.3)添加集合数据  {key:value}  key=动态列的key值
            for (int x = 0; x < 5; x++) {
                contentMap.put(sortKey + x, "数据" + x);
            }
            //3.4)集合添加到数组中
            dataList.add(contentMap);
        }
        //4、定义大标题
        ExportParams ex = new ExportParams(title, "sheet1");
        //导出工具
        ExportUtil.dynamicExport(title,resourceEntityList,ex,dataList,null,false);
    }
}
