package excel.exceldownload.web;

import excel.exceldownload.domain.ExcelDownloadService;
import excel.exceldownload.domain.dto.FormatSampleExcelDto;
import excel.exceldownload.excel.CustomSXSSFExcelFile;
import excel.exceldownload.excel.SXSSFExcelFile;
import excel.exceldownload.domain.dto.SampleExcelDto;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
@RequiredArgsConstructor
public class ExcelDownloadController {

    private final ExcelDownloadService excelDownloadService;

    @GetMapping("/api/excel/v1")
    public void downloadExcelV1(HttpServletResponse response) throws IOException {
        excelDownloadService.downloadExcel(response);
    }

    @GetMapping("/api/excel/v2")
    public void downloadExcelV2(HttpServletResponse response) throws IOException {
        List<SampleExcelDto> samples = new ArrayList<>();
        samples.add(new SampleExcelDto(0L, "1번 이름", "1번 제목"));
        samples.add(new SampleExcelDto(1L, "2번 이름", "2번 제목"));

        SXSSFExcelFile<SampleExcelDto> excelFile = new SXSSFExcelFile<>(samples, SampleExcelDto.class);
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=example.xlsx");
        excelFile.write(response.getOutputStream());
    }

    @GetMapping("/api/excel/v3")
    public void downloadExcelV3(HttpServletResponse response) throws IOException {
        excelDownloadService.downloadFormatExcel(response);
    }

    @GetMapping("/api/excel/v4")
    public void downloadExcelV4(HttpServletResponse response) throws IOException {
        List<FormatSampleExcelDto> samples = new ArrayList<>();
        samples.add(new FormatSampleExcelDto("기업규모", "대기업", "Count", 45, 42, 87));
        samples.add(new FormatSampleExcelDto("기업규모", "대기업", "% Within 기업 규모", 45, 42, 87));
        samples.add(new FormatSampleExcelDto("기업규모", "중견.중소기업", "Count", 45, 42, 87));
        samples.add(new FormatSampleExcelDto("기업규모", "중견.중소기업", "% Within 기업 규모", 45, 42, 87));
        samples.add(new FormatSampleExcelDto("수출비중", "수출기업", "Count", 45, 42, 87));
        samples.add(new FormatSampleExcelDto("수출비중", "수출기업", "% Within 기업 규모", 45, 42, 87));
        samples.add(new FormatSampleExcelDto("수출비중", "내수기업", "Count", 45, 42, 87));
        samples.add(new FormatSampleExcelDto("수출비중", "내수기업", "% Within 기업 규모", 45, 42, 87));

        SXSSFExcelFile<FormatSampleExcelDto> excelFile = new CustomSXSSFExcelFile<>("타이틀", 2, 2, samples, FormatSampleExcelDto.class);
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=example.xlsx");
        excelFile.write(response.getOutputStream());
    }

}
