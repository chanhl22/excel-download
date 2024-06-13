package excel.exceldownload.sample;

import lombok.Builder;
import lombok.Getter;

import java.util.List;

@Getter
public class Statistics {

    private final List<String> statisticsColumnQuestions;

    private final List<StatisticsExcelData> statisticsExcelDatas;

    @Builder
    private Statistics(List<String> statisticsColumnQuestions, List<StatisticsExcelData> statisticsExcelDatas) {
        this.statisticsColumnQuestions = statisticsColumnQuestions;
        this.statisticsExcelDatas = statisticsExcelDatas;
    }

    public static Statistics createStatisticsQuestion(List<String> statisticsColumnQuestions, List<StatisticsExcelData> statisticsExcelDatas) {
        return Statistics.builder()
                .statisticsColumnQuestions(statisticsColumnQuestions)
                .statisticsExcelDatas(statisticsExcelDatas)
                .build();
    }

}
