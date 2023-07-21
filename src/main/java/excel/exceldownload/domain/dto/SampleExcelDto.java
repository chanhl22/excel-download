package excel.exceldownload.domain.dto;

import excel.exceldownload.annotation.ExcelColumn;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

@NoArgsConstructor
@AllArgsConstructor
public class SampleExcelDto {

    @ExcelColumn(headerName = "번호")
    private long id;

    @ExcelColumn(headerName = "이름")
    private String name;

    @ExcelColumn(headerName = "제목")
    private String title;
}
