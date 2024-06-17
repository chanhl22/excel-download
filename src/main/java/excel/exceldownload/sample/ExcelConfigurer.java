package excel.exceldownload.sample;

import excel.exceldownload.sample.alignment.ExcelAlignment;
import excel.exceldownload.sample.borders.ExcelMergedRegionBorders;
import excel.exceldownload.sample.font.ExcelFont;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

@Component
@RequiredArgsConstructor
public class ExcelConfigurer {

    private final ExcelRenderResourceFactory excelRenderResourceFactory;

    public void configure(Workbook workbook, Sheet sheet, int firstRow, int endRow, int firstCol, int endCol, String cellValue,
                          ExcelMergedRegionBorders excelMergedRegionBorders, ExcelFont excelFont, ExcelAlignment excelAlignment) {

        generateCellValue(sheet, firstRow, firstCol, cellValue);
        generateFont(workbook, sheet, firstRow, firstCol, excelFont, excelAlignment);
        generateMergedRegion(sheet, firstRow, endRow, firstCol, endCol);
        generateBorders(sheet, firstRow, endRow, firstCol, endCol, excelMergedRegionBorders);
    }

    public void configure(Workbook workbook, Sheet sheet, int firstRow, int firstCol, String cellValue,
                          ExcelMergedRegionBorders excelMergedRegionBorders, ExcelFont excelFont, ExcelAlignment excelAlignment) {

        generateCellValue(sheet, firstRow, firstCol, cellValue);
        generateFont(workbook, sheet, firstRow, firstCol, excelFont, excelAlignment);
        generateBorders(sheet, firstRow, firstRow, firstCol, firstCol, excelMergedRegionBorders);
    }

    private void generateCellValue(Sheet sheet, int rowNum, int cellNum, String cellValue) {
        excelRenderResourceFactory.setCellValue(sheet, rowNum, cellNum, cellValue);
    }

    private void generateFont(Workbook workbook, Sheet sheet, int rowNum, int cellNum, ExcelFont excelFont, ExcelAlignment excelAlignment) {
        excelRenderResourceFactory.applyFont(workbook, sheet, rowNum, cellNum, excelFont, excelAlignment);
    }

    private void generateMergedRegion(Sheet sheet, int firstRow, int endRow, int firstCol, int endCol) {
        excelRenderResourceFactory.addMergedRegion(sheet, firstRow, endRow, firstCol, endCol);
    }

    private void generateBorders(Sheet sheet, int firstRow, int endRow, int firstCol, int endCol, ExcelMergedRegionBorders excelMergedRegionBorders) {
        excelRenderResourceFactory.applyMergedRegionBorders(sheet, firstRow, endRow, firstCol, endCol, excelMergedRegionBorders);
    }

}
