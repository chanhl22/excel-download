package excel.exceldownload.domain;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Service
public class ExcelDownloadService {

    public void downloadExcel(HttpServletResponse response) throws IOException {
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("1번 시트");
        Row row;
        Cell cell;
        int rowNum = 0;

        // Header
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("번호");
        cell = row.createCell(1);
        cell.setCellValue("이름");
        cell = row.createCell(2);
        cell.setCellValue("제목");

        // Body
        for (int i = 0; i < 3; i++) {
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue(i);
            cell = row.createCell(1);
            cell.setCellValue(i + "_name");
            cell = row.createCell(2);
            cell.setCellValue(i + "_title");
        }

        // 컨텐츠 타입과 파일명 지정
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=example.xlsx");

        // Excel File Output
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    public void downloadFormatExcel(HttpServletResponse response) throws IOException {
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("1번 시트");
        Row row;
        Cell cell;
        CellRangeAddress mergedRange;

        // 0.
        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("귀사는 혁신기업이라고 생각하십니까?");
        row.createCell(1);
        row.createCell(2);
        row.createCell(3);
        row.createCell(4);
        row.createCell(5);
        mergedRange = new CellRangeAddress(0, 0, 0, 5);
        sheet.addMergedRegion(mergedRange);

        // 1.
        row = sheet.createRow(1);
        row.createCell(0);
        row.createCell(1);
        row.createCell(2);
        cell = row.createCell(3);
        cell.setCellValue("그렇다");
        cell = row.createCell(4);
        cell.setCellValue("아니다");
        cell = row.createCell(5);
        cell.setCellValue("Total");
        mergedRange = new CellRangeAddress(1, 1, 0, 2);
        sheet.addMergedRegion(mergedRange);

        // 2. 3.
        Row row2 = sheet.createRow(2);

        Cell row2cell0 = row2.createCell(0);
        row2cell0.setCellValue("Total");

        row2.createCell(1);

        Cell row2cell2 = row2.createCell(2);
        row2cell2.setCellValue("Count");

        Cell row2cell3 = row2.createCell(3);
        row2cell3.setCellValue(150);

        Cell row2cell4 = row2.createCell(4);
        row2cell4.setCellValue(150);

        Cell row2cell5 = row2.createCell(5);
        row2cell5.setCellValue(300);


        Row row3 = sheet.createRow(3);

        row3.createCell(0);

        row3.createCell(1);

        Cell row3cell2 = row3.createCell(2);
        row3cell2.setCellValue("% within 기업규모");

        Cell row3cell3 = row3.createCell(3);
        row3cell3.setCellValue("50.0%");

        Cell row3cell4 = row3.createCell(4);
        row3cell4.setCellValue("50.0%");

        Cell row3cell5 = row3.createCell(5);
        row3cell5.setCellValue("100.0%");

        mergedRange = new CellRangeAddress(2, 3, 0, 1);
        sheet.addMergedRegion(mergedRange);

        // 4. 5. 6. 7.
        Row row4 = sheet.createRow(4);
        Row row5 = sheet.createRow(5);
        Row row6 = sheet.createRow(6);
        Row row7 = sheet.createRow(7);

        Cell row4cell0 = row4.createCell(0);
        row4cell0.setCellValue("기업규모");

        Cell row4cell1 = row4.createCell(1);
        row4cell1.setCellValue("대기업");

        Cell row4cell2 = row4.createCell(2);
        row4cell2.setCellValue("Count");

        Cell row4cell3 = row4.createCell(3);
        row4cell3.setCellValue(45);

        Cell row4cell4 = row4.createCell(4);
        row4cell4.setCellValue(42);

        Cell row4cell5 = row4.createCell(5);
        row4cell5.setCellValue(87);


        row5.createCell(0);
        row5.createCell(1);

        Cell row5cell2 = row5.createCell(2);
        row5cell2.setCellValue("% within 기업규모");

        Cell row5cell3 = row5.createCell(3);
        row5cell3.setCellValue("51.7%");

        Cell row5cell4 = row5.createCell(4);
        row5cell4.setCellValue("48.3%");

        Cell row5cell5 = row5.createCell(5);
        row5cell5.setCellValue("100.0%");


        row6.createCell(0);
        Cell row6cell1 = row6.createCell(1);
        row6cell1.setCellValue("중견.중소기업");

        Cell row6cell2 = row6.createCell(2);
        row6cell2.setCellValue("Count");

        Cell row6cell3 = row6.createCell(3);
        row6cell3.setCellValue(105);

        Cell row6cell4 = row6.createCell(4);
        row6cell4.setCellValue(108);

        Cell row6cell5 = row6.createCell(5);
        row6cell5.setCellValue(213);


        row7.createCell(0);
        row7.createCell(1);

        Cell row7cell2 = row7.createCell(2);
        row7cell2.setCellValue("% within 기업규모");

        Cell row7cell3 = row7.createCell(3);
        row7cell3.setCellValue("49.3%");

        Cell row7cell4 = row7.createCell(4);
        row7cell4.setCellValue("50.7%%");

        Cell row7cell5 = row7.createCell(5);
        row7cell5.setCellValue("100.0%");


        mergedRange = new CellRangeAddress(4, 7, 0, 0);
        sheet.addMergedRegion(mergedRange);

        mergedRange = new CellRangeAddress(4, 5, 1, 1);
        sheet.addMergedRegion(mergedRange);
        mergedRange = new CellRangeAddress(6, 7, 1, 1);
        sheet.addMergedRegion(mergedRange);


        // 컨텐츠 타입과 파일명 지정
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=example.xlsx");

        // Excel File Output
        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
