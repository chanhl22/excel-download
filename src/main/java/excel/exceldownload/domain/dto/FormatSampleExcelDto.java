package excel.exceldownload.domain.dto;

import excel.exceldownload.main.annotation.DefaultBodyStyle;
import excel.exceldownload.main.annotation.DefaultHeaderStyle;
import excel.exceldownload.main.annotation.ExcelColumn;
import excel.exceldownload.main.annotation.ExcelColumnStyle;
import excel.exceldownload.main.excel.style.custom.AlignCenterAndBordersThinBodyStyle;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@NoArgsConstructor
@AllArgsConstructor
@DefaultHeaderStyle(style = @ExcelColumnStyle(excelCellStyleClass = AlignCenterAndBordersThinBodyStyle.class))
@DefaultBodyStyle(style = @ExcelColumnStyle(excelCellStyleClass = AlignCenterAndBordersThinBodyStyle.class))
public class FormatSampleExcelDto {

    @ExcelColumn(headerName = " ")
    private String parent;

    @ExcelColumn(headerName = " ")
    private String child;

    @ExcelColumn(headerName = " ")
    private String type;

    @ExcelColumn(
            headerName = "그렇다",
            columnSize = 20 * 256 + 100
    )
    private long yes;

    @ExcelColumn(
            headerName = "아니다",
            columnSize = 20 * 256 + 100
    )
    private long no;

    @ExcelColumn(
            headerName = "Total",
            columnSize = 20 * 256 + 100
    )
    private long total;

}
