package excel.exceldownload.dynamic.dto;

import excel.exceldownload.dynamic.annotation.ExcelBody;
import excel.exceldownload.dynamic.annotation.ExcelColumnStyle;
import excel.exceldownload.dynamic.annotation.ExcelHeader;
import excel.exceldownload.dynamic.style.custom.AlignCenterAndBordersThinBodyStyle;
import excel.exceldownload.dynamic.style.custom.LightGreyHeaderStyle;
import lombok.Builder;
import lombok.Getter;

import java.util.List;
import java.util.Map;

@Getter
public class SurveyResultDetailExcelDownloadDto {

    @ExcelHeader(
            columnSize = 10000,
            style = @ExcelColumnStyle(excelCellStyleClass = LightGreyHeaderStyle.class)
    )
    private final List<String> headers;

    @ExcelBody(
            style = @ExcelColumnStyle(excelCellStyleClass = AlignCenterAndBordersThinBodyStyle.class)
    )
    private final List<Map<String, Object>> bodies;

    @Builder(builderMethodName = "innerBuilder")
    private SurveyResultDetailExcelDownloadDto(List<String> headers, List<Map<String, Object>> bodies) {
        this.headers = headers;
        this.bodies = bodies;
    }

    public static SurveyResultDetailExcelDownloadDtoBuilder builder(List<String> headers, List<Map<String, Object>> bodies) {
        return innerBuilder()
                .headers(headers)
                .bodies(bodies);
    }

}
