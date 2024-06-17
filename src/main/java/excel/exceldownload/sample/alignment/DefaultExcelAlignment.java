package excel.exceldownload.sample.alignment;

import lombok.Builder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class DefaultExcelAlignment implements ExcelAlignment {

    private final HorizontalAlignment alignment;
    private final VerticalAlignment verticalAlignment;

    @Builder
    private DefaultExcelAlignment(HorizontalAlignment alignment, VerticalAlignment verticalAlignment) {
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

    public static DefaultExcelAlignment create() {
        return DefaultExcelAlignment.builder()
                .alignment(HorizontalAlignment.GENERAL)
                .verticalAlignment(VerticalAlignment.CENTER)
                .build();
    }

}
