package excel.exceldownload.sample.alignment;

import org.apache.poi.ss.usermodel.CellStyle;

public interface ExcelAlignment {

    void applyAlignment(CellStyle cellStyle);

    void applyVerticalAlignment(CellStyle cellStyle);

    void apply(CellStyle cellStyle);

}
