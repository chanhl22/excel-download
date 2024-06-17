package excel.exceldownload.asis.sample;

import lombok.Builder;
import lombok.Getter;

@Getter
public class MainQuestion {

    private final String questionMessage;

    private final int firstRowIndex;

    private final int firstColumnIndex;

    private final int lastRowIndex;

    private final int lastColumnIndex;

    @Builder
    private MainQuestion(String questionMessage, int firstRowIndex, int firstColumnIndex, int lastRowIndex, int lastColumnIndex) {
        this.questionMessage = questionMessage;
        this.firstRowIndex = firstRowIndex;
        this.firstColumnIndex = firstColumnIndex;
        this.lastRowIndex = lastRowIndex;
        this.lastColumnIndex = lastColumnIndex;
    }

    public static MainQuestion createMainQuestion(String questionMessage, int startRowIndex, int startColumnIndex, int lastRowIndex, int lastColumnIndex) {
        return MainQuestion.builder()
                .questionMessage(questionMessage)
                .firstRowIndex(startRowIndex)
                .firstColumnIndex(startColumnIndex)
                .lastRowIndex(lastRowIndex)
                .lastColumnIndex(lastColumnIndex)
                .build();
    }

}
