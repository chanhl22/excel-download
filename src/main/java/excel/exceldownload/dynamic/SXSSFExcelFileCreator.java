package excel.exceldownload.dynamic;

import excel.exceldownload.dynamic.annotation.ExcelBody;
import excel.exceldownload.dynamic.annotation.ExcelHeader;
import excel.exceldownload.dynamic.resource.ExcelRenderResource;
import excel.exceldownload.dynamic.resource.ExcelRenderResourceFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;

public class SXSSFExcelFileCreator<T> {

    private static final int ROW_START_INDEX = 0;
    private static final int COLUMN_START_INDEX = 0;
    private int currentRowIndex = ROW_START_INDEX;

    protected SXSSFWorkbook workbook;
    protected Sheet sheet;
    protected ExcelRenderResource resource;

    public SXSSFExcelFileCreator(T data, Class<T> type) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException, InstantiationException {
        this.workbook = new SXSSFWorkbook();
        this.resource = ExcelRenderResourceFactory.prepareRenderResource(type, workbook);
        renderExcel(data, type);
    }

    protected void renderExcel(T data, Class<T> clazz) throws IllegalAccessException {
        sheet = workbook.createSheet();
        Field[] fields = clazz.getDeclaredFields();
        List<?> headers = List.of();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelHeader.class)) {
                field.setAccessible(true);
                List<?> list = (List<?>) field.get(data);
                String fieldName = field.getName();
                renderHeadersWithNewSheet(list, fieldName, currentRowIndex++);
                headers = list;
            }
        }

        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelBody.class)) {
                field.setAccessible(true);
                List<?> bodies = (List<?>) field.get(data);
                String fieldName = field.getName();
                for (Object body : bodies) {
                    renderBody(headers, (Map<?, ?>) body, fieldName, currentRowIndex++);
                }
            }
        }
    }

    protected void renderHeadersWithNewSheet(List<?> headers, String fieldName, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = SXSSFExcelFileCreator.COLUMN_START_INDEX;
        for (Object header : headers) {
            sheet.setColumnWidth(columnIndex, resource.getExcelColumnSize(fieldName));
            Cell cell = row.createCell(columnIndex++);
            cell.setCellStyle(resource.getCellStyle(fieldName));
            renderCellValue(cell, header);
        }
    }

    protected void renderBody(List<?> headers, Map<?, ?> bodies, String fieldName, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        int columnIndex = SXSSFExcelFileCreator.COLUMN_START_INDEX;
        for (Object header : headers) {
            Cell cell = row.createCell(columnIndex++);
            cell.setCellStyle(resource.getCellStyle(fieldName));
            renderCellValue(cell, bodies.get(header));
        }
    }

    protected void renderCellValue(Cell cell, Object cellValue) {
        if (cellValue instanceof LocalDateTime) {
            LocalDateTime localDataTimeValue = (LocalDateTime) cellValue;
            String formatted = localDataTimeValue.format(DateTimeFormatter.ofPattern("yyyy년 MM월 dd일"));
            cell.setCellValue(formatted);
            return;
        }

        if (cellValue instanceof Number) {
            Number numberValue = (Number) cellValue;
            cell.setCellValue(numberValue.doubleValue());
            return;
        }

        cell.setCellValue(cellValue == null ? "-" : getCustomString(cellValue));
    }

    protected static String getCustomString(Object cellValue) {
        String value = cellValue.toString();
        return value.replace("||", "\n").replaceAll("<[^>]+>", "");
    }

    public void write(OutputStream stream) throws IOException {
        workbook.write(stream);
        workbook.close();
        workbook.dispose();
        stream.close();
    }

}
