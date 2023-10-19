package com.muyangren.excel.style;

import cn.afterturn.easypoi.excel.export.styler.ExcelExportStylerDefaultImpl;
import org.apache.poi.ss.usermodel.*;

/**
 * @author: muyangren
 * @Date: 2023/2/14
 * @Description: 继承默认ExcelExportStylerDefaultImpl() 并按自己需求进行重写
 * @Version: 1.0
 */
public class ExcelExportStyler extends ExcelExportStylerDefaultImpl {
    public ExcelExportStyler(Workbook workbook) {
        super(workbook);
    }

    @Override
    public CellStyle getHeaderStyle(short color) {
        //初始化CellStyle
        CellStyle titleStyle = initCellStyle();
        //填充表格颜色
        titleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //增加样式
        titleStyle.setFont(getFont(workbook, (short) 50, true));
        return titleStyle;
    }


    @Override
    public CellStyle getTitleStyle(short color) {
        //初始化CellStyle
        CellStyle titleStyle = initCellStyle();
        titleStyle.setWrapText(true);
        //增加字体样式
        titleStyle.setFont(getFont(workbook, (short) 35, true));
        return titleStyle;
    }

    public CellStyle initCellStyle() {
        CellStyle titleStyle = this.workbook.createCellStyle();
        //加粗
        //下边框
        titleStyle.setBorderBottom(BorderStyle.THIN);
        //左边框
        titleStyle.setBorderLeft(BorderStyle.THIN);
        //上边框
        titleStyle.setBorderTop(BorderStyle.THIN);
        //右边框
        titleStyle.setBorderRight(BorderStyle.THIN);
        //居中
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return titleStyle;
    }

    private Font getFont(Workbook workbook, short size, boolean isBold) {
        Font font = workbook.createFont();
        //字体样式
        font.setFontName("宋体");
        //是否加粗
        font.setBold(isBold);
        //字体大小
        font.setFontHeightInPoints(size);
        return font;
    }
}
