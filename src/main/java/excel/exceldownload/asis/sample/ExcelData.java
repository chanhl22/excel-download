package excel.exceldownload.asis.sample;

import lombok.Builder;
import lombok.Getter;

@Getter
public class ExcelData {

    private final String parentTag;

    private final String childTag;

    private final String columnQuestion;

    private final String rowQuestion;

    private final Object data;

    @Builder
    private ExcelData(String parentTag, String childTag, String columnQuestion, String rowQuestion, Object data) {
        this.parentTag = parentTag;
        this.childTag = childTag;
        this.columnQuestion = columnQuestion;
        this.rowQuestion = rowQuestion;
        this.data = data;
    }

}
