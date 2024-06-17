package excel.exceldownload.asis.sample;

import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@RestController
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService service;

    @GetMapping("/api/admin/v1/excel")
    public void downloadExcelV1(HttpServletResponse response) throws IOException {
        service.sample(response);
    }

}
