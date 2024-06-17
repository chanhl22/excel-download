package excel.exceldownload.sample.alignment;

import lombok.Builder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class CustomExcelAlignment implements ExcelAlignment {

    private final HorizontalAlignment alignment;
    private final VerticalAlignment verticalAlignment;

    @Builder
    private CustomExcelAlignment(HorizontalAlignment alignment, VerticalAlignment verticalAlignment) {
        this.alignment = alignment;
        this.verticalAlignment = verticalAlignment;
    }

    @Override
    public void applyAlignment(CellStyle cellStyle) {
        cellStyle.setAlignment(alignment);
    }

    @Override
    public void applyVerticalAlignment(CellStyle cellStyle) {
        cellStyle.setVerticalAlignment(verticalAlignment);
    }

    @Override
    public void apply(CellStyle cellStyle) {
        cellStyle.setAlignment(alignment);
        cellStyle.setVerticalAlignment(verticalAlignment);
    }

    public static CustomExcelAlignment create(HorizontalAlignment alignment, VerticalAlignment verticalAlignment) {
        return CustomExcelAlignment.builder()
                .alignment(alignment)
                .verticalAlignment(verticalAlignment)
                .build();
    }

}
