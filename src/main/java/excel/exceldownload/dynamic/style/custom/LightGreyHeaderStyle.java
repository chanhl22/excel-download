package excel.exceldownload.dynamic.style.custom;

import excel.exceldownload.dynamic.style.CustomExcelCellStyle;
import excel.exceldownload.dynamic.style.align.DefaultExcelAlign;
import excel.exceldownload.dynamic.style.border.DefaultExcelBorders;
import excel.exceldownload.dynamic.style.border.ExcelBorderStyle;
import excel.exceldownload.dynamic.style.configurer.ExcelCellStyleConfigurer;

public class LightGreyHeaderStyle extends CustomExcelCellStyle {

    @Override
    public void configure(ExcelCellStyleConfigurer configurer) {
        configurer
                .excelAlign(DefaultExcelAlign.CENTER_CENTER)
                .foregroundColor(211, 211, 211)
                .excelBorders(DefaultExcelBorders.newInstance(ExcelBorderStyle.THIN));
    }

}