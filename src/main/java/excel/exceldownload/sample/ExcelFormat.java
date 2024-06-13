package excel.exceldownload.sample;

import lombok.Builder;
import lombok.Getter;

import java.util.List;
import java.util.stream.Collectors;

@Getter
public class ExcelFormat {

    private final int startRowIndex;

    private final int startColumnIndex;


    private final int startRowQuestionCellIndex;

    private final int startColumnQuestionCellIndex;

    private final int statisticsRowStartPosition;

    private final int statisticsColumnStartPosition;

    private final List<String> parentTags;

    private final List<String> childTags;

    private final List<String> rowQuestions;

    private final List<String> columnQuestions;

    private final List<ExcelData> excelDatas;

    private final List<String> statisticsColumnQuestions;

    private final List<StatisticsExcelData> statisticsExcelDatas;

    private final Statistics statistics;

    private final MainQuestion mainQuestion;

    @Builder
    private ExcelFormat(String questionMessage, List<ExcelData> excelDatas, List<StatisticsExcelData> statisticsExcelDatas, int startRowIndex, int startColumnIndex) {
        this.startRowIndex = startRowIndex; // 시작 row
        this.startColumnIndex = startColumnIndex; // 시작 cell

        this.startRowQuestionCellIndex = startRowIndex + 3; // 질문의 cell 위치 (그렇다, 아니다)
        this.startColumnQuestionCellIndex = startColumnIndex + 2; //질문의 cell 위치 (count, 기업규모 위치)

        this.statisticsRowStartPosition = startRowIndex + 1; //빈칸 첫번째 row
        this.statisticsColumnStartPosition = startColumnIndex; //빈칸 첫번째 cell

        this.parentTags = excelDatas.stream()
                .map(ExcelData::getParentTag)
                .distinct()
                .collect(Collectors.toList());
        this.childTags = excelDatas.stream()
                .map(ExcelData::getChildTag)
                .distinct()
                .collect(Collectors.toList());
        this.rowQuestions = excelDatas.stream()
                .map(ExcelData::getRowQuestion)
                .distinct()
                .collect(Collectors.toList());
        this.columnQuestions = excelDatas.stream()
                .map(ExcelData::getColumnQuestion)
                .distinct()
                .collect(Collectors.toList());
        this.excelDatas = excelDatas;
        this.statisticsColumnQuestions = statisticsExcelDatas.stream()
                .map(StatisticsExcelData::getColumnQuestion)
                .distinct()
                .collect(Collectors.toList());
        this.statisticsExcelDatas = statisticsExcelDatas;


        this.statistics = Statistics.createStatisticsQuestion(
                extractStatisticsQuestions(statisticsExcelDatas),
                statisticsExcelDatas);
        this.mainQuestion = MainQuestion.createMainQuestion(questionMessage, startRowIndex, startColumnIndex, startRowIndex, startRowQuestionCellIndex + rowQuestions.size() - 1);
    }

    private List<String> extractStatisticsQuestions(List<StatisticsExcelData> statisticsExcelDatas) {
        return statisticsExcelDatas.stream()
                .map(StatisticsExcelData::getColumnQuestion)
                .distinct()
                .collect(Collectors.toList());
    }

    public int getLastColumnCount() {
        return startRowQuestionCellIndex + rowQuestions.size();
    }

}
