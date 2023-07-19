package excel.exceldownload.web;

import excel.exceldownload.domain.ExcelDownloadService;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@Controller
@RequiredArgsConstructor
public class ExcelDownloadController {

    private final ExcelDownloadService excelDownloadService;

    @GetMapping("/api/excel")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        excelDownloadService.downloadExcel(response);
    }
}
