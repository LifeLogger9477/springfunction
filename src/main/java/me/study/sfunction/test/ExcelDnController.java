package me.study.sfunction.test;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.IOException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * author : ms.Lee
 * date   : 2024-01-02
 */

@Controller
public class ExcelDnController {

  @GetMapping (value = "/")
  public String home() {

    return "home";
  }

  @GetMapping (value = "/excel/download")
  public void excelDownload(HttpServletResponse response) throws IOException {

    Workbook wb = new XSSFWorkbook();
    Sheet sheet = wb.createSheet( "1번 시트" );
    Row row = null;
    Cell cell = null;

    int numRow = 0;

    // Header
    row = sheet.createRow( numRow++ );
    cell = row.createCell( 0 );
    cell.setCellValue( "번호" );

    cell = row.createCell( 1 );
    cell.setCellValue( "이름" );

    cell = row.createCell( 2 );
    cell.setCellValue( "제목" );

    // Body
    for (int i = 0; i < 3; i++) {

      row = sheet.createRow( numRow++ );
      cell = row.createCell( 0 );
      cell.setCellValue( i );

      cell = row.createCell( 1 );
      cell.setCellValue( "name_" + i );

      cell = row.createCell( 2 );
      cell.setCellValue( "title_" + i );
    }

    // content type 과 파일명 지정
    // 현재 시간
    Date now = new Date();
    SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
    String fileName = "엑셀파일명_" + sdf.format( now ) + ".xlsx";
    fileName = URLEncoder.encode( fileName, "UTF-8" );
    response.setContentType( "application/vnd.msexcel" );
    response.setHeader(
        "Content-Disposition",
        "attachment; filename=\"" + fileName + "\""
    );

    // Excel File Output
    wb.write( response.getOutputStream() );
    wb.close();
  }
}
