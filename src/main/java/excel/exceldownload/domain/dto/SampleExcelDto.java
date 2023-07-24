package excel.exceldownload.domain.dto;

import excel.exceldownload.annotation.DefaultBodyStyle;
import excel.exceldownload.annotation.DefaultHeaderStyle;
import excel.exceldownload.annotation.ExcelColumn;
import excel.exceldownload.annotation.ExcelColumnStyle;
import excel.exceldownload.excel.style.NoExcelCellStyle;
import excel.exceldownload.excel.style.custom.BlueHeaderStyle;
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
            bodyStyle = @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class)
    )
    private long id;

    @ExcelColumn(headerName = "이름")
    private String name;

    @ExcelColumn(headerName = "제목")
    private String title;
}
