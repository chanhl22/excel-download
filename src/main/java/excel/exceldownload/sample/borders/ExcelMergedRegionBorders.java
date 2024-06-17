package excel.exceldownload.sample.borders;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public interface ExcelMergedRegionBorders {

    void applyBorders(Sheet sheet, CellRangeAddress region);

    void applyBorderColor(Sheet sheet, CellRangeAddress region);

    void apply(Sheet sheet, CellRangeAddress region);

}
