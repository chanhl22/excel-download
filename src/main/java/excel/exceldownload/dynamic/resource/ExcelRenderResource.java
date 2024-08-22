package excel.exceldownload.dynamic.resource;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Map;

public class ExcelRenderResource {

    private final Map<String, Integer> excelColumnSizes;
    private final PreCalculatedCellStyleMap styleMap;

    public ExcelRenderResource(Map<String, Integer> excelColumnSizes, PreCalculatedCellStyleMap styleMap) {
        this.excelColumnSizes = excelColumnSizes;
        this.styleMap = styleMap;
    }

    public Integer getExcelColumnSize(String dataFieldName) {
        return excelColumnSizes.get(dataFieldName);
    }

    public CellStyle getCellStyle(String dataFieldName) {
        return styleMap.get(dataFieldName);
    }

}
