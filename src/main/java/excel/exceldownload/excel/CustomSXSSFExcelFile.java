package excel.exceldownload.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class CustomSXSSFExcelFile<T> extends SXSSFExcelFile<T> {

    private final String title;

    public CustomSXSSFExcelFile(String title, int rowStart, int columnStart, List<T> data, Class<T> type) {
        super(type, rowStart + 1, columnStart);
        this.title = title;
        renderExcel(data);
    }

    @Override
    protected void renderExcel(List<T> data) {
        // 1. Create sheet and renderHeader
        sheet = workbook.createSheet();
        renderTitle(sheet, currentRowIndex - 1);
        renderHeadersWithNewSheet(sheet, currentRowIndex++);

        if (data.isEmpty()) {
            return;
        }

        // 2. Render Body
        for (Object renderedData : data) {
            renderBody(renderedData, currentRowIndex++);
        }

        mergeBody();
    }

    private void renderTitle(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = startColumnIndex;

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);

        for (String ignored : resource.getDataFieldNames()) {
            Cell cell = row.createCell(columnIndex++);
            cell.setCellStyle(cellStyle);
            renderCellValue(cell, title);
        }

        CellRangeAddress cellRange = new CellRangeAddress(rowIndex, rowIndex, startColumnIndex, columnIndex - 1);
        sheet.addMergedRegion(cellRange);
    }

    private void mergeBody() {
        for (int columnIndex = startColumnIndex; columnIndex < startColumnIndex + 2; columnIndex++) {
            int standardRowIndex = sheet.getFirstRowNum() + 2;
            for (int rowIndex = sheet.getFirstRowNum() + 2; rowIndex < sheet.getLastRowNum(); rowIndex++) {
                Row row1 = sheet.getRow(rowIndex);
                Row row2 = sheet.getRow(rowIndex + 1);
                Cell cell1 = row1.getCell(columnIndex);
                Cell cell2 = row2.getCell(columnIndex);
                if (!cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    sheet.addMergedRegion(new CellRangeAddress(standardRowIndex, rowIndex, columnIndex, columnIndex));
                    standardRowIndex = rowIndex + 1;
                }

                if (rowIndex == sheet.getLastRowNum() - 1) {
                    sheet.addMergedRegion(new CellRangeAddress(standardRowIndex, sheet.getLastRowNum(), columnIndex, columnIndex));
                }
            }
        }
    }



}
