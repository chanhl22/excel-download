package excel.exceldownload.main.excel.resource.collection;

import excel.exceldownload.main.excel.resource.ExcelCellKey;
import excel.exceldownload.main.excel.style.ExcelCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class PreCalculatedCellStyleMap {

    private final Map<ExcelCellKey, CellStyle> cellStyleMap = new HashMap<>();

    public PreCalculatedCellStyleMap() {
    }

    public void put(ExcelCellKey excelCellKey, ExcelCellStyle excelCellStyle, Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        excelCellStyle.apply(cellStyle);
        cellStyleMap.put(excelCellKey, cellStyle);
    }

    public CellStyle get(ExcelCellKey excelCellKey) {
        return cellStyleMap.get(excelCellKey);
    }

    public boolean isEmpty() {
        return cellStyleMap.isEmpty();
    }
}
