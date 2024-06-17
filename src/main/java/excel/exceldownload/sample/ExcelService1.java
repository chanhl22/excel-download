package excel.exceldownload.sample;

import excel.exceldownload.sample.alignment.CustomExcelAlignment;
import excel.exceldownload.sample.alignment.DefaultExcelAlignment;
import excel.exceldownload.sample.borders.DefaultMergedRegionBorders;
import excel.exceldownload.sample.borders.NoMergedRegionBorders;
import excel.exceldownload.sample.font.CustomExcelFont;
import excel.exceldownload.sample.font.DefaultExcelFont;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Service
@RequiredArgsConstructor
public class ExcelService1 {

    private final ExcelSheetFormatBuilder excelSheetFormatBuilder;
    private final ExcelConfigurer excelConfigurer;

    public void downloadExcel(HttpServletResponse response) throws IOException {
        Workbook workbook = new SXSSFWorkbook();

        Sheet sheet = excelSheetFormatBuilder.generateExcelSheetFormat(workbook, "Sheet1", 0, 0, 12, 6);

        excelConfigurer.configure(workbook, sheet, 0, 0, 0, 5, "1.귀사는 혁신 기업이라고 생각하십니까?",
                DefaultMergedRegionBorders.create(),
                CustomExcelFont.create("sample", true, Font.U_NONE, HSSFColor.HSSFColorPredefined.BLACK),
                DefaultExcelAlignment.create());

        excelConfigurer.configure(workbook, sheet, 4, 7, 0, 0, "기업규모",
                NoMergedRegionBorders.create(),
                DefaultExcelFont.create(),
                CustomExcelAlignment.create(HorizontalAlignment.CENTER, VerticalAlignment.CENTER));

        excelConfigurer.configure(workbook, sheet, 4, 5, 1, 1, "대기업",
                NoMergedRegionBorders.create(),
                DefaultExcelFont.create(),
                CustomExcelAlignment.create(HorizontalAlignment.CENTER, VerticalAlignment.CENTER));

        excelConfigurer.configure(workbook, sheet, 6, 7, 1, 1, "중견.중소기업",
                NoMergedRegionBorders.create(),
                DefaultExcelFont.create(),
                CustomExcelAlignment.create(HorizontalAlignment.CENTER, VerticalAlignment.CENTER));

        excelConfigurer.configure(workbook, sheet, 4, 2, "Count",
                NoMergedRegionBorders.create(),
                DefaultExcelFont.create(),
                CustomExcelAlignment.create(HorizontalAlignment.CENTER, VerticalAlignment.CENTER));

        excelConfigurer.configure(workbook, sheet, 4, 3, "45",
                NoMergedRegionBorders.create(),
                DefaultExcelFont.create(),
                CustomExcelAlignment.create(HorizontalAlignment.CENTER, VerticalAlignment.CENTER));

        // 컨텐츠 타입과 파일명 지정
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=example.xlsx");

        // Excel File Output
        workbook.write(response.getOutputStream());
        workbook.close();
    }


}
