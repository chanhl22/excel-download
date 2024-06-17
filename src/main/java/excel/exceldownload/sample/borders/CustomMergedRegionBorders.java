package excel.exceldownload.sample.borders;

import lombok.Builder;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

public class CustomMergedRegionBorders implements ExcelMergedRegionBorders {

    private final BorderStyle borderStyle;

    private final IndexedColors borderColor;

    @Builder
    private CustomMergedRegionBorders(BorderStyle borderStyle, IndexedColors borderColor) {
        this.borderStyle = borderStyle;
        this.borderColor = borderColor;
    }

    @Override
    public void applyBorders(Sheet sheet, CellRangeAddress region) {
        RegionUtil.setBorderTop(borderStyle, region, sheet);
        RegionUtil.setBorderBottom(borderStyle, region, sheet);
        RegionUtil.setBorderLeft(borderStyle, region, sheet);
        RegionUtil.setBorderRight(borderStyle, region, sheet);
    }

    @Override
    public void applyBorderColor(Sheet sheet, CellRangeAddress region) {
        RegionUtil.setTopBorderColor(borderColor.index, region, sheet);
        RegionUtil.setBottomBorderColor(borderColor.index, region, sheet);
        RegionUtil.setLeftBorderColor(borderColor.index, region, sheet);
        RegionUtil.setRightBorderColor(borderColor.index, region, sheet);
    }

    @Override
    public void apply(Sheet sheet, CellRangeAddress region) {
        RegionUtil.setBorderTop(borderStyle, region, sheet);
        RegionUtil.setBorderBottom(borderStyle, region, sheet);
        RegionUtil.setBorderLeft(borderStyle, region, sheet);
        RegionUtil.setBorderRight(borderStyle, region, sheet);

        RegionUtil.setTopBorderColor(borderColor.index, region, sheet);
        RegionUtil.setBottomBorderColor(borderColor.index, region, sheet);
        RegionUtil.setLeftBorderColor(borderColor.index, region, sheet);
        RegionUtil.setRightBorderColor(borderColor.index, region, sheet);
    }

    public static CustomMergedRegionBorders create(BorderStyle borderStyle, IndexedColors borderColor) {
        return CustomMergedRegionBorders.builder()
                .borderStyle(borderStyle)
                .borderColor(borderColor)
                .build();
    }

}
