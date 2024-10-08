package excel.exceldownload.main.excel;

import excel.exceldownload.main.excel.resource.ExcelRenderLocation;
import excel.exceldownload.main.excel.resource.ExcelRenderResource;
import excel.exceldownload.main.excel.resource.ExcelRenderResourceFactory;
import excel.exceldownload.main.exception.ExcelInternalException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

import static excel.exceldownload.main.utils.ReflectionUtils.getField;

public class SXSSFExcelFile<T> {

    private static final int ROW_START_INDEX = 0;
    private static final int COLUMN_START_INDEX = 0;
    private int currentRowIndex = ROW_START_INDEX;

    protected SXSSFWorkbook workbook;
    protected Sheet sheet;
    protected ExcelRenderResource resource;

    public SXSSFExcelFile(List<T> data, Class<T> type) {
        this.workbook = new SXSSFWorkbook();
        this.resource = ExcelRenderResourceFactory.prepareRenderResource(type, workbook);
        renderExcel(data);
    }

    public SXSSFExcelFile(Class<T> type) {
        this.workbook = new SXSSFWorkbook();
        this.resource = ExcelRenderResourceFactory.prepareRenderResource(type, workbook);
    }

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
    }

    protected void renderHeadersWithNewSheet(int rowIndex, int startColumnIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = startColumnIndex;
        for (String dataFieldName : resource.getDataFieldNames()) {
            sheet.setColumnWidth(columnIndex, resource.getExcelColumnSize(dataFieldName));
            Cell cell = row.createCell(columnIndex++);
            cell.setCellStyle(resource.getCellStyle(dataFieldName, ExcelRenderLocation.HEADER));
            cell.setCellValue(resource.getExcelHeaderName(dataFieldName));
        }
    }

    protected void renderBody(Object data, int rowIndex, int startColumnIndex) {
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
