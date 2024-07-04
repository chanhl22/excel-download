package excel.exceldownload.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class CustomSXSSFExcelFile<T> extends SXSSFExcelFile<T> {

    private final String title;

    public CustomSXSSFExcelFile(String title, List<T> data, Class<T> type) {
        super(type);
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
    }

    private void renderTitle(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = COLUMN_START_INDEX;

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

        CellRangeAddress cellRange = new CellRangeAddress(rowIndex, rowIndex, COLUMN_START_INDEX, columnIndex - 1);
        sheet.addMergedRegion(cellRange);
    }

}
