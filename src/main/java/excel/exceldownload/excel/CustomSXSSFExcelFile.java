package excel.exceldownload.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class CustomSXSSFExcelFile<T> extends SXSSFExcelFile<T> {

    private final int ROW_START_INDEX;
    private final int COLUMN_START_INDEX;
    private int currentRowIndex;
    private final String title;

    public CustomSXSSFExcelFile(String title, int rowStart, int columnStart, List<T> data, Class<T> type) {
        super(type);
        this.ROW_START_INDEX = rowStart;
        this.COLUMN_START_INDEX = columnStart;
        this.currentRowIndex = rowStart + 1;
        this.title = title;
        renderExcel(data);
    }

    @Override
    protected void renderExcel(List<T> data) {
        // 1. Create sheet and renderHeader
        sheet = workbook.createSheet();
        renderHeadersWithNewSheet(currentRowIndex++, COLUMN_START_INDEX);

        if (data.isEmpty()) {
            return;
        }

        // 2. Render Body
        for (Object renderedData : data) {
            renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);
        }

        // 3. Render Title
        renderTitle();

        // 4. Merge Body
        mergeBody();
    }

    private void renderTitle() {
        Row row = sheet.createRow(ROW_START_INDEX);
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

        CellRangeAddress cellRange = new CellRangeAddress(ROW_START_INDEX, ROW_START_INDEX, COLUMN_START_INDEX, columnIndex - 1);
        sheet.addMergedRegion(cellRange);
    }

    private void mergeBody() {
        int startRowNum = sheet.getFirstRowNum() + 2;
        int lastRowNum = sheet.getLastRowNum();
        int startColNum = sheet.getRow(startRowNum).getFirstCellNum();
        int lastColNum = startColNum + 3;

        for (int col = startColNum; col < lastColNum; col++) {
            String prevValue = null;
            int startRow = -1;

            for (int row = startRowNum; row <= lastRowNum; row++) {
                Row currentRow = sheet.getRow(row);
                if (currentRow != null) {
                    Cell currentCell = currentRow.getCell(col);
                    String currentValue = currentCell != null ? currentCell.getStringCellValue() : "";

                    if (!currentValue.equals(prevValue)) {
                        if (startRow != -1 && startRow != row - 1) {
                            sheet.addMergedRegion(new CellRangeAddress(startRow, row - 1, col, col));
                        }
                        startRow = row;
                        prevValue = currentValue;
                    }
                }
            }

            if (startRow != -1 && startRow != lastRowNum) {
                sheet.addMergedRegion(new CellRangeAddress(startRow, lastRowNum, col, col));
            }
        }
    }

}
