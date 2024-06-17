package excel.exceldownload.asis.sample;

import lombok.Builder;
import lombok.Getter;

@Getter
public class StatisticsExcelData {

    private final String columnQuestion;

    private final String rowQuestion;

    private final Object data;

    @Builder
    private StatisticsExcelData(String columnQuestion, String rowQuestion, Object data) {
        this.columnQuestion = columnQuestion;
        this.rowQuestion = rowQuestion;
        this.data = data;
    }

}
