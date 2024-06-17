package excel.exceldownload.sample;

import excel.exceldownload.sample.alignment.ExcelAlignment;
import excel.exceldownload.sample.borders.ExcelMergedRegionBorders;
import excel.exceldownload.sample.font.ExcelFont;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;

@Component
public class ExcelRenderResourceFactory {

    public void setCellValue(Sheet sheet, int rowNum, int cellNum, String value) {
        sheet.getRow(rowNum).getCell(cellNum).setCellValue(value);
    }

    public void addMergedRegion(Sheet sheet, int firstRow, int endRow, int firstCol, int endCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, endRow, firstCol, endCol));
    }

    public void applyMergedRegionBorders(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, ExcelMergedRegionBorders excelMergedRegionBorders) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        excelMergedRegionBorders.apply(sheet, region);
    }

    public void applyFont(Workbook workbook, Sheet sheet, int rowNum, int cellNum, ExcelFont excelFont, ExcelAlignment excelAlignment) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();

        excelFont.apply(style, font);
        excelAlignment.apply(style);

        sheet.getRow(rowNum).getCell(cellNum).setCellStyle(style);
    }

}
