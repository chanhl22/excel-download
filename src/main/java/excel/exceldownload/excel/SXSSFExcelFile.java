package excel.exceldownload.excel;

import excel.exceldownload.excel.resource.ExcelRenderLocation;
import excel.exceldownload.excel.resource.ExcelRenderResource;
import excel.exceldownload.excel.resource.ExcelRenderResourceFactory;
import excel.exceldownload.exception.ExcelInternalException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

import static excel.exceldownload.utils.ReflectionUtils.getField;

public class SXSSFExcelFile<T> {

    protected static final int ROW_START_INDEX = 0;
    protected static final int COLUMN_START_INDEX = 0;
    protected int currentRowIndex = ROW_START_INDEX;
    protected int startColumnIndex = COLUMN_START_INDEX;

    protected SXSSFWorkbook workbook;
    protected Sheet sheet;
    protected ExcelRenderResource resource;

    public SXSSFExcelFile(List<T> data, Class<T> type) {
        this.workbook = new SXSSFWorkbook();
        this.resource = ExcelRenderResourceFactory.prepareRenderResource(type, workbook);
        renderExcel(data);
    }

    public SXSSFExcelFile(Class<T> type, int rowStart, int columnStart) {
        currentRowIndex = rowStart;
        startColumnIndex = columnStart;
        this.workbook = new SXSSFWorkbook();
        this.resource = ExcelRenderResourceFactory.prepareRenderResource(type, workbook);
    }

    protected void renderExcel(List<T> data) {
        // 1. Create sheet and renderHeader
        sheet = workbook.createSheet();
        renderHeadersWithNewSheet(sheet, currentRowIndex++);

        if (data.isEmpty()) {
            return;
        }

        // 2. Render Body
        for (Object renderedData : data) {
            renderBody(renderedData, currentRowIndex++);
        }
    }

    protected void renderHeadersWithNewSheet(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = startColumnIndex;
        for (String dataFieldName : resource.getDataFieldNames()) {
            sheet.setColumnWidth(columnIndex, resource.getExcelColumnSize(dataFieldName));
            Cell cell = row.createCell(columnIndex++);
            cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.HEADER));
            cell.setCellValue(resource.getExcelHeaderName(dataFieldName));
        }
    }

    protected void renderBody(Object data, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = startColumnIndex;
        for (String dataFieldName : resource.getDataFieldNames()) {
            Cell cell = row.createCell(columnIndex++);
            try {
                Field field = getField(data.getClass(), dataFieldName);
                field.setAccessible(true);
                cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.BODY));
                Object cellValue = field.get(data);
                renderCellValue(cell, cellValue);
            } catch (Exception e) {
                throw new ExcelInternalException(e.getMessage(), e);
            }
        }
    }

    protected void renderCellValue(Cell cell, Object cellValue) {
        if (cellValue instanceof Number) {
            Number numberValue = (Number) cellValue;
            cell.setCellValue(numberValue.doubleValue());
            return;
        }
        cell.setCellValue(cellValue == null ? "" : cellValue.toString());
    }

    public void write(OutputStream stream) throws IOException {
        workbook.write(stream);
        workbook.close();
        workbook.dispose();
        stream.close();
    }
}
