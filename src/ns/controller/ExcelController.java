package ns.controller;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("/excel")
public class ExcelController {

    @Value("${webBaseUrl}")
    private String webBaseUrl; // 실제 다운로드 URL 만들 때 사용

    @RequestMapping("/excelDownloadFdt.do")
    public void excelDownloadFdt(HttpServletResponse response) throws IOException  {

        String templateUrl = webBaseUrl + "/template/LAPTOP_REPORT_TEMPLATE.xlsx";

        // 1. URL로부터 InputStream 얻기
        URL url = new URL(templateUrl);
        InputStream is = url.openStream();
        
        // POI로 가공
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        if (row == null) row = sheet.createRow(0);
        Cell cell = row.getCell(1);
        if (cell == null) cell = row.createCell(1);
        cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()));

        // 응답 전송
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"result.xlsx\"");

        OutputStream os = response.getOutputStream();
        workbook.write(os);
        os.flush();

        workbook.close();
        is.close();
    }
    
    @RequestMapping("/excelDownloadFdtEntity.do")
    public ResponseEntity<byte[]> excelDownloadFdtEntity() throws IOException {

        String templateUrl = webBaseUrl + "/template/LAPTOP_REPORT_TEMPLATE.xlsx";

        // 1. URL로부터 InputStream 얻기
        URL url = new URL(templateUrl);
        InputStream is = url.openStream();

        // 2. POI로 가공
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        if (row == null) row = sheet.createRow(0);
        Cell cell = row.getCell(1);
        if (cell == null) cell = row.createCell(1);
        cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()));

        // 3. ByteArrayOutputStream에 쓰기
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);

        // 4. 자원 정리
        workbook.close();
        is.close();

        // 5. ResponseEntity 생성
        byte[] content = bos.toByteArray();

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "result.xlsx");

        return new ResponseEntity<>(content, headers, HttpStatus.OK);
    } 
}
