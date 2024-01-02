package me.study.sfunction.test;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
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

    // 테이블 헤더용
    CellStyle headerStyle = wb.createCellStyle();

    // 가는 경계선
    headerStyle.setBorderTop( BorderStyle.THIN );
    headerStyle.setBorderBottom( BorderStyle.THIN );
    headerStyle.setBorderLeft( BorderStyle.THIN );
    headerStyle.setBorderRight( BorderStyle.THIN );

    // 배경색
    headerStyle.setFillForegroundColor(
        HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex()
    );
    headerStyle.setFillPattern( FillPatternType.SOLID_FOREGROUND );

    // 데이터용 경계선
    CellStyle bodyStyle = wb.createCellStyle();
    bodyStyle.setBorderTop( BorderStyle.DASHED );
    bodyStyle.setBorderBottom( BorderStyle.DASHED );
    bodyStyle.setBorderLeft( BorderStyle.DASHED );
    bodyStyle.setBorderRight( BorderStyle.DASHED );

    // Header
    String[] headerArray = { "NO.", "제목", "내용", "등록일", "등록자", "사용여부" };
    row = sheet.createRow( numRow++ );
    for (int col = 0; col < headerArray.length; col++) {

      cell = row.createCell( col );
      cell.setCellStyle( headerStyle );
      cell.setCellValue( headerArray[col] );
    }

    // Body
    for (int i = 0; i < 3; i++) {

      row = sheet.createRow( numRow++ );
      cell = row.createCell( 0 );
      cell.setCellStyle( bodyStyle );
      cell.setCellValue( i );

      cell = row.createCell( 1 );
      cell.setCellStyle( bodyStyle );
      cell.setCellValue( "name_" + i );

      cell = row.createCell( 2 );
      cell.setCellStyle( bodyStyle );
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
