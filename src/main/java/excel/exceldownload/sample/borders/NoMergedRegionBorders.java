package excel.exceldownload.sample.borders;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class NoMergedRegionBorders implements ExcelMergedRegionBorders {

    public NoMergedRegionBorders() {
    }

    @Override
    public void applyBorders(Sheet sheet, CellRangeAddress region) {

    }

    @Override
    public void applyBorderColor(Sheet sheet, CellRangeAddress region) {

    }

    @Override
    public void apply(Sheet sheet, CellRangeAddress region) {

    }

    public static NoMergedRegionBorders create() {
        return new NoMergedRegionBorders();
    }

}
