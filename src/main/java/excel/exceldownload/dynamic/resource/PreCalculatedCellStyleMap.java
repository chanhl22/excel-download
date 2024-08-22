package excel.exceldownload.dynamic.resource;

import excel.exceldownload.dynamic.style.ExcelCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class PreCalculatedCellStyleMap {

    private final Map<String, CellStyle> cellStyleMap = new HashMap<>();

    public PreCalculatedCellStyleMap() {
    }

    public void put(String excelCellKey, ExcelCellStyle excelCellStyle, Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(true);
        excelCellStyle.apply(cellStyle);
        cellStyleMap.put(excelCellKey, cellStyle);
    }

    public CellStyle get(String excelCellKey) {
        return cellStyleMap.get(excelCellKey);
    }

    public boolean isEmpty() {
        return cellStyleMap.isEmpty();
    }

}
