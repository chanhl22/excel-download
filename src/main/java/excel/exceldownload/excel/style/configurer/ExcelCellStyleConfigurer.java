package excel.exceldownload.excel.style.configurer;

import excel.exceldownload.excel.style.align.ExcelAlign;
import excel.exceldownload.excel.style.align.NoExcelAlign;
import excel.exceldownload.excel.style.border.ExcelBorders;
import excel.exceldownload.excel.style.border.NoExcelBorders;
import excel.exceldownload.excel.style.color.DefaultExcelColor;
import excel.exceldownload.excel.style.color.ExcelColor;
import excel.exceldownload.excel.style.color.NoExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelCellStyleConfigurer {

    private ExcelAlign excelAlign = new NoExcelAlign();
    private ExcelColor foregroundColor = new NoExcelColor();
    private ExcelBorders excelBorders = new NoExcelBorders();

    public ExcelCellStyleConfigurer excelAlign(ExcelAlign excelAlign) {
        this.excelAlign = excelAlign;
        return this;
    }

    public ExcelCellStyleConfigurer foregroundColor(int red, int blue, int green) {
        this.foregroundColor = DefaultExcelColor.rgb(red, blue, green);
        return this;
    }

    public ExcelCellStyleConfigurer excelBorders(ExcelBorders excelBorders) {
        this.excelBorders = excelBorders;
        return this;
    }

    public void configure(CellStyle cellStyle) {
        excelAlign.apply(cellStyle);
        foregroundColor.applyForeground(cellStyle);
        excelBorders.apply(cellStyle);
    }
}