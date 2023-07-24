package excel.exceldownload.excel.style.custom;

import excel.exceldownload.excel.style.CustomExcelCellStyle;
import excel.exceldownload.excel.style.border.DefaultExcelBorders;
import excel.exceldownload.excel.style.border.ExcelBorderStyle;
import excel.exceldownload.excel.style.configurer.ExcelCellStyleConfigurer;

public class BlueHeaderStyle extends CustomExcelCellStyle {

    @Override
    public void configure(ExcelCellStyleConfigurer configurer) {
        configurer.foregroundColor(223, 235, 246)
                .excelBorders(DefaultExcelBorders.newInstance(ExcelBorderStyle.THIN))
                .excelAlign(com.lannstark.style.align.DefaultExcelAlign.CENTER_CENTER);
    }
}