package excel.exceldownload.sample.font;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

public interface ExcelFont {

    void applyFontName(CellStyle cellStyle, Font font);

    void applyBold(CellStyle cellStyle, Font font);

    void applyUnderline(CellStyle cellStyle, Font font);

    void applyColor(CellStyle cellStyle, Font font);

    void apply(CellStyle cellStyle, Font font);

}
