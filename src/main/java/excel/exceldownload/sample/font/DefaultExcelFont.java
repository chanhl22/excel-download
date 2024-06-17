package excel.exceldownload.sample.font;

import lombok.Builder;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

public class DefaultExcelFont implements ExcelFont {

    private final String fontName;
    private final boolean bold;
    private final byte underline;
    private final HSSFColor.HSSFColorPredefined color;

    @Builder
    private DefaultExcelFont(String fontName, boolean bold, byte underline, HSSFColor.HSSFColorPredefined color) {
        this.fontName = fontName;
        this.bold = bold;
        this.underline = underline;
        this.color = color;
    }

    @Override
    public void applyFontName(CellStyle cellStyle, Font font) {
        font.setFontName(fontName);
        cellStyle.setFont(font);
    }

    @Override
    public void applyBold(CellStyle cellStyle, Font font) {
        font.setBold(bold);
        cellStyle.setFont(font);
    }

    @Override
    public void applyUnderline(CellStyle cellStyle, Font font) {
        font.setUnderline(underline);
        cellStyle.setFont(font);
    }

    @Override
    public void applyColor(CellStyle cellStyle, Font font) {
        font.setColor(color.getIndex());
        cellStyle.setFont(font);
    }

    @Override
    public void apply(CellStyle cellStyle, Font font) {
        font.setFontName(fontName);
        font.setBold(bold);
        font.setUnderline(underline);
        font.setColor(color.getIndex());
        cellStyle.setFont(font);
    }

    public static DefaultExcelFont create() {
        return DefaultExcelFont.builder()
                .fontName("default")
                .bold(false)
                .underline(Font.U_NONE)
                .color(HSSFColor.HSSFColorPredefined.BLACK)
                .build();
    }

}
