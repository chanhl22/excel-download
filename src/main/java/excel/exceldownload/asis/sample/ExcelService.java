package excel.exceldownload.asis.sample;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;


import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class ExcelService {

    public void sample(HttpServletResponse response) throws IOException {
        downloadFormatExcel(response, ExcelFormat.builder()
                .questionMessage("귀사는 혁신기업이라고 생각하십니까?")
                .excelDatas(List.of(ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("대기업")
                                .columnQuestion("Count")
                                .rowQuestion("그렇다")
                                .data(45)
                                .build(),
                        ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("대기업")
                                .columnQuestion("Count")
                                .rowQuestion("아니다")
                                .data(42)
                                .build(),
                        ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("대기업")
                                .columnQuestion("% within 기업규모")
                                .rowQuestion("그렇다")
                                .data("51.7%")
                                .build(),
                        ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("대기업")
                                .columnQuestion("% within 기업규모")
                                .rowQuestion("아니다")
                                .data("48.3%")
                                .build(),
                        ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("중견.중소기업")
                                .columnQuestion("Count")
                                .rowQuestion("그렇다")
                                .data(105)
                                .build(),
                        ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("중견.중소기업")
                                .columnQuestion("Count")
                                .rowQuestion("아니다")
                                .data(108)
                                .build(),
                        ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("중견.중소기업")
                                .columnQuestion("% within 기업규모")
                                .rowQuestion("그렇다")
                                .data("49.3%")
                                .build(),
                        ExcelData.builder()
                                .parentTag("기업규모")
                                .childTag("중견.중소기업")
                                .columnQuestion("% within 기업규모")
                                .rowQuestion("아니다")
                                .data("50.7%")
                                .build()
                ))
                .statisticsExcelDatas(List.of(StatisticsExcelData.builder()
                                .columnQuestion("Count")
                                .rowQuestion("그렇다")
                                .data(150)
                                .build(),
                        StatisticsExcelData.builder()
                                .columnQuestion("Count")
                                .rowQuestion("아니다")
                                .data(150)
                                .build(),
                        StatisticsExcelData.builder()
                                .columnQuestion("% within 기업규모")
                                .rowQuestion("그렇다")
                                .data("50.0%")
                                .build(),
                        StatisticsExcelData.builder()
                                .columnQuestion("% within 기업규모")
                                .rowQuestion("아니다")
                                .data("50.0%")
                                .build()
                ))
                        .startRowIndex(0)
                        .startColumnIndex(0)
                        .build());
    }

    public void downloadFormatExcel(HttpServletResponse response, ExcelFormat excelFormat) throws IOException {
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("1번 시트");
        AtomicInteger rowIndex = new AtomicInteger(excelFormat.getStartRowIndex());

        createMainQuestionFormat(sheet, excelFormat, rowIndex);

        //
        createStatisticsFormat(sheet, excelFormat, rowIndex);

        //
        createMain(sheet, excelFormat, rowIndex);


        // 컨텐츠 타입과 파일명 지정
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=example.xlsx");

        // Excel File Output
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    private void createMainQuestionFormat(Sheet sheet, ExcelFormat excelFormat, AtomicInteger rowIndex) {
        int currentRowIndex = rowIndex.getAndIncrement();
        MainQuestion mainQuestion = excelFormat.getMainQuestion();

        Row row = sheet.createRow(currentRowIndex);
        for (int i = mainQuestion.getFirstColumnIndex(); i < mainQuestion.getLastColumnIndex(); i++) {
            row.createCell(i);
        }

        sheet.addMergedRegion(new CellRangeAddress(
                mainQuestion.getFirstRowIndex(),
                mainQuestion.getLastRowIndex(),
                mainQuestion.getFirstColumnIndex(),
                mainQuestion.getLastColumnIndex()));

        row.getCell(mainQuestion.getFirstRowIndex())
                .setCellValue(mainQuestion.getQuestionMessage());
    }

    private void createMain(Sheet sheet, ExcelFormat excelFormat, AtomicInteger rowIndex) {
        // 시작 row 인덱스
        int startRowIndex = rowIndex.get();

        // 총 만들어질 row 수를 구하기
        // 1. 부모 태그 수
        // 2. 자식 태그 수
        // 3. 세로 질문 수
        int parentTagCount = excelFormat.getParentTags().size();
        int childTagCount = excelFormat.getChildTags().size();
        int columnQuestionCount = excelFormat.getColumnQuestions().size();

        //총 row
        int totalRowCount = parentTagCount * childTagCount * columnQuestionCount;

        //첫번째 column 인덱스
        int firstColumn = excelFormat.getStartColumnIndex();

        // 모든 cell 생성
        for (int i = 0; i < totalRowCount; i++) {
            Row row = sheet.createRow(rowIndex.getAndIncrement());
            for (int j = firstColumn; j < firstColumn + excelFormat.getLastColumnCount(); j++) {
                row.createCell(j);
            }
        }

        // parent 태그 채우기
        int parentRowCount = childTagCount * columnQuestionCount;
        int start = startRowIndex;
        for (int j = 0; j < parentTagCount; j++) {
            for (int i = start; i < start + parentRowCount; i++) {
                sheet.getRow(i).getCell(excelFormat.getStartColumnIndex()).setCellValue(excelFormat.getParentTags().get(j));
            }
            start += start + parentRowCount;
        }

//        int parentRowCount = childTagCount * columnQuestionCount;
//        for (int i = startRowIndex, j = 0; j < parentTagCount; i += parentRowCount, j++) {
//
//            sheet.getRow(i).getCell(excelFormat.getStartColumnIndex()).setCellValue(excelFormat.getParentTags().get(j));
//        }

        // child 태그 채우기
        Map<String, List<ExcelData>> parentTagMap = excelFormat.getExcelDatas().stream()
                .collect(Collectors.groupingBy(ExcelData::getParentTag));


        for (int i = 0; i < excelFormat.getParentTags().size(); i++) {
            String parentTag = excelFormat.getParentTags().get(i);
            List<ExcelData> parentData = parentTagMap.get(parentTag);
            List<String> childTagList = parentData.stream()
                    .map(ExcelData::getChildTag)
                    .distinct()
                    .collect(Collectors.toList());

            int childRowCount = columnQuestionCount;
            int start1 = startRowIndex;
            for (int j = 0; j < childTagList.size(); j++) {
                for (int l = start1; l < start1 + childRowCount; l++) {
                    sheet.getRow(l).getCell(excelFormat.getStartColumnIndex() + 1).setCellValue(childTagList.get(j));
                }
                start1 += childRowCount;
            }
        }


        // question 채우기
        for (int i = 0; i < excelFormat.getParentTags().size(); i++) {
            String parentTag = excelFormat.getParentTags().get(i);
            int y = startRowIndex;
            for (int j = 0; j < excelFormat.getChildTags().size(); j++) {
                String childTag = excelFormat.getChildTags().get(j);
                Map<String, List<ExcelData>> childTagMap = parentTagMap.get(parentTag).stream()
                        .collect(Collectors.groupingBy(ExcelData::getChildTag));
                List<ExcelData> childData = childTagMap.get(childTag);

                List<String> columnQuestions = childData.stream()
                        .map(ExcelData::getColumnQuestion)
                        .distinct()
                        .collect(Collectors.toList());

                int count = 0;

                for (int k = y, x = 0; x < columnQuestions.size(); k ++, x++) {
                    sheet.getRow(k).getCell(excelFormat.getStartColumnIndex() + 2).setCellValue(columnQuestions.get(x));
                    count++;
                }
                y += count;
            }
        }

        //
        // Main 값 채우기
        // 시작 ROW
        for (int i = startRowIndex; i < startRowIndex + totalRowCount; i++) {
            int startColumn = excelFormat.getStartRowQuestionCellIndex();
            int endColumn = excelFormat.getStartRowQuestionCellIndex() + excelFormat.getColumnQuestions().size();

            int columnValueIndex1 = startColumn - 3;
            int columnValueIndex2 = startColumn - 2;
            int columnValueIndex3 = startColumn - 1;

            for (int j = startColumn; j < endColumn; j++) {
                String columnValue1 = sheet.getRow(i).getCell(columnValueIndex1).getStringCellValue();
                String columnValue2 = sheet.getRow(i).getCell(columnValueIndex2).getStringCellValue();
                String columnValue3 = sheet.getRow(i).getCell(columnValueIndex3).getStringCellValue();
                String rowValue = sheet.getRow(excelFormat.getStatisticsRowStartPosition()).getCell(j).getStringCellValue();

                List<ExcelData> excelDatas = excelFormat.getExcelDatas();


                sheet.getRow(i).getCell(j).setCellValue(
                excelDatas.stream()
                        .filter(excelData -> excelData.getParentTag().equals(columnValue1) && excelData.getChildTag().equals(columnValue2) && excelData.getColumnQuestion().equals(columnValue3) && excelData.getRowQuestion().equals(rowValue))
                        .map(excelData -> String.valueOf(excelData.getData()))
                        .findFirst()
                        .orElse(""));


            }

        }


    }

    private void createStatisticsFormat(Sheet sheet, ExcelFormat excelFormat, AtomicInteger rowIndex) {
        createStatisticsRow(sheet, rowIndex, excelFormat);
        createStatisticsColumn(sheet, rowIndex, excelFormat);
        createStatisticsData(sheet, rowIndex, excelFormat);
    }

    private void createStatisticsRow(Sheet sheet, AtomicInteger rowIndex, ExcelFormat excelFormat) {
        int currentRowIndex = rowIndex.getAndIncrement();
        Statistics statistics = excelFormat.getStatistics();


        int currentColumnIndex = excelFormat.getStartColumnIndex();

        Row row = sheet.createRow(currentRowIndex);
        for (int i = currentColumnIndex; i < excelFormat.getLastColumnCount(); i++) {
            row.createCell(i);
        }

        AtomicInteger rowQuestionIndex = new AtomicInteger(0);
        excelFormat.getRowQuestions()
                .forEach(rowQuestion -> row.getCell(currentColumnIndex + excelFormat.getStartRowQuestionCellIndex() + rowQuestionIndex.getAndIncrement())
                        .setCellValue(rowQuestion));

        sheet.addMergedRegion(new CellRangeAddress(currentRowIndex, currentRowIndex, currentColumnIndex, currentColumnIndex + excelFormat.getStartRowQuestionCellIndex() - 1));
    }

    private void createStatisticsColumn(Sheet sheet, AtomicInteger rowIndex, ExcelFormat excelFormat) {
        int currentColumnIndex = excelFormat.getStartColumnIndex();

        excelFormat.getStatisticsColumnQuestions()
                .forEach(columnQuestion -> {
                    Row row = sheet.createRow(rowIndex.getAndIncrement());
                    for (int i = currentColumnIndex; i < excelFormat.getLastColumnCount(); i++) {
                        row.createCell(i);
                    }
                    row.getCell(excelFormat.getStartColumnQuestionCellIndex()).setCellValue(columnQuestion);
                });

        sheet.getRow(excelFormat.getStatisticsRowStartPosition() + 1)
                .getCell(excelFormat.getStatisticsColumnStartPosition())
                .setCellValue("Total");

        sheet.addMergedRegion(new CellRangeAddress(excelFormat.getStatisticsRowStartPosition() + 1,
                excelFormat.getStatisticsRowStartPosition() + 2,
                excelFormat.getStatisticsColumnStartPosition(),
                excelFormat.getStatisticsColumnStartPosition() + 1));
    }

    private void createStatisticsData(Sheet sheet, AtomicInteger rowIndex, ExcelFormat excelFormat) {
        Map<String, List<StatisticsExcelData>> collect = excelFormat.getStatisticsExcelDatas().stream()
                .collect(Collectors.groupingBy(StatisticsExcelData::getColumnQuestion));

        int startRow = excelFormat.getStatisticsRowStartPosition() + 1;
        int endRow = excelFormat.getStatisticsRowStartPosition() + 1 + excelFormat.getRowQuestions().size();
        int startColumn = excelFormat.getStartRowQuestionCellIndex();
        int endColumn = excelFormat.getStartRowQuestionCellIndex() + excelFormat.getColumnQuestions().size();

        for (int i = startRow; i < endRow; i++) {
            for (int j = startColumn, k = startColumn - 1; j < endColumn; j++) {
                String columnValue = sheet.getRow(i).getCell(k).getStringCellValue();
                String rowValue = sheet.getRow(excelFormat.getStatisticsRowStartPosition()).getCell(j).getStringCellValue();
                sheet.getRow(i).getCell(j).setCellValue(
                        collect.get(columnValue).stream()
                                .filter(statisticsExcelData -> statisticsExcelData.getRowQuestion().equals(rowValue))
                                .map(statisticsExcelData -> String.valueOf(statisticsExcelData.getData()))
                                .findFirst()
                                .orElse(""));
            }
        }
    }

}
