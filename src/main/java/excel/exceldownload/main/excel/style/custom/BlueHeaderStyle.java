package excel.exceldownload.main.excel.style.custom;

import excel.exceldownload.main.excel.style.CustomExcelCellStyle;
import excel.exceldownload.main.excel.style.align.DefaultExcelAlign;
import excel.exceldownload.main.excel.style.border.DefaultExcelBorders;
import excel.exceldownload.main.excel.style.border.ExcelBorderStyle;
import excel.exceldownload.main.excel.style.configurer.ExcelCellStyleConfigurer;

public class BlueHeaderStyle extends CustomExcelCellStyle {

    @Override
    public void configure(ExcelCellStyleConfigurer configurer) {
        configurer
                .excelAlign(DefaultExcelAlign.CENTER_CENTER)
                .foregroundColor(223, 235, 246)
                .excelBorders(DefaultExcelBorders.newInstance(ExcelBorderStyle.THIN));
    }
}