package excel.exceldownload.excel.resource;

import excel.exceldownload.excel.resource.collection.PreCalculatedCellStyleMap;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;
import java.util.Map;

public class ExcelRenderResource {

    private PreCalculatedCellStyleMap styleMap;
    private Map<String, String> excelHeaderNames;
    private Map<String, Integer> excelColumnSizes;
    private List<String> dataFieldNames;

    public ExcelRenderResource(PreCalculatedCellStyleMap styleMap,
                               Map<String, String> excelHeaderNames, Map<String, Integer> excelColumnSizes, List<String> dataFieldNames) {
        this.styleMap = styleMap;
        this.excelHeaderNames = excelHeaderNames;
        this.dataFieldNames = dataFieldNames;
        this.excelColumnSizes = excelColumnSizes;
    }

    public CellStyle getCellStyle(String dataFieldName, ExcelRenderLocation excelRenderLocation) {
        return styleMap.get(ExcelCellKey.of(dataFieldName, excelRenderLocation));
    }

    public String getExcelHeaderName(String dataFieldName) {
        return excelHeaderNames.get(dataFieldName);
    }

    public Integer getExcelColumnSize(String dataFieldName) {
        return excelColumnSizes.get(dataFieldName);
    }

    public List<String> getDataFieldNames() {
        return dataFieldNames;
    }
}
