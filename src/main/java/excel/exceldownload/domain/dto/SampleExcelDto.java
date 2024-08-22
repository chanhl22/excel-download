package excel.exceldownload.domain.dto;

import excel.exceldownload.main.annotation.DefaultBodyStyle;
import excel.exceldownload.main.annotation.DefaultHeaderStyle;
import excel.exceldownload.main.annotation.ExcelColumn;
import excel.exceldownload.main.annotation.ExcelColumnStyle;
import excel.exceldownload.main.excel.style.NoExcelCellStyle;
import excel.exceldownload.main.excel.style.custom.AlignCenterAndBordersThinBodyStyle;
import excel.exceldownload.main.excel.style.custom.BlueHeaderStyle;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

@NoArgsConstructor
@AllArgsConstructor
@DefaultHeaderStyle(style = @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class))
@DefaultBodyStyle(style = @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class))
public class SampleExcelDto {

    @ExcelColumn(
            headerName = "번호",
            headerStyle = @ExcelColumnStyle(excelCellStyleClass = BlueHeaderStyle.class),
            bodyStyle = @ExcelColumnStyle(excelCellStyleClass = AlignCenterAndBordersThinBodyStyle.class)
    )
    private long id;

    @ExcelColumn(
            headerName = "이름",
            columnSize = 20 * 256 + 100,
            headerStyle = @ExcelColumnStyle(excelCellStyleClass = BlueHeaderStyle.class)
    )
    private String name;

    @ExcelColumn(headerName = "제목")
    private String title;
}
