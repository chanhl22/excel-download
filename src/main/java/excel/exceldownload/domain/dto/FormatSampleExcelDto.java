package excel.exceldownload.domain.dto;

import excel.exceldownload.annotation.DefaultBodyStyle;
import excel.exceldownload.annotation.DefaultHeaderStyle;
import excel.exceldownload.annotation.ExcelColumn;
import excel.exceldownload.annotation.ExcelColumnStyle;
import excel.exceldownload.excel.style.NoExcelCellStyle;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Getter
@NoArgsConstructor
@AllArgsConstructor
@DefaultHeaderStyle(style = @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class))
@DefaultBodyStyle(style = @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class))
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
