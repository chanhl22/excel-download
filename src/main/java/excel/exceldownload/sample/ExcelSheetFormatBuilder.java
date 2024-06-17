package excel.exceldownload.sample;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.util.stream.IntStream;

@Component
public class ExcelSheetFormatBuilder {

    public Sheet generateExcelSheetFormat(Workbook workbook, String sheetName, int startRow, int startColumn, int rowSize, int cellSize) {
        Sheet sheet = workbook.createSheet(sheetName);
        generateRows(startRow, startColumn, rowSize, cellSize, sheet);
        return sheet;
    }

    private void generateRows(int startRow, int startColumn, int rowSize, int cellSize, Sheet sheet) {
        IntStream.range(startRow, startRow + rowSize)
                .forEach(rowIndex -> {
                    Row row = sheet.createRow(rowIndex);
                    generateCells(startColumn, cellSize, row);
                });
    }

    private void generateCells(int startColumn, int cellSize, Row row) {
        IntStream.range(startColumn, startColumn + cellSize)
                .forEach(row::createCell);
    }

}
