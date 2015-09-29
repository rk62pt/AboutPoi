package com.ryan.xls.web;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.ryan.xls.web.vo.Book;

import org.apache.poi.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletResponse;

@Controller
public class XlsController {
	@RequestMapping(value = "index", method = RequestMethod.GET)
	public String index() {
		
		return "index.html";
	}
	
	@RequestMapping(value = "/downloadXLS")
    public void downloadCSV(HttpServletResponse response) throws IOException {
 
//        String csvFileName = "books.csv";
// 
//        response.setContentType("appl");
 
        Workbook workbook = new XSSFWorkbook();;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        //字體格式
        Font font = workbook.createFont();
        font.setColor(IndexedColors.BLACK.getIndex()); // 顏色
        font.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
        // 設定儲存格格式 
        CellStyle styleRow1 = workbook.createCellStyle();
        // styleRow1.setFillForegroundColor(HSSFColor.GREEN.index);//填滿顏色
        // styleRow1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        styleRow1.setFont(font); // 設定字體
        styleRow1.setAlignment(CellStyle.ALIGN_CENTER); // 水平置中
        styleRow1.setVerticalAlignment(CellStyle.VERTICAL_CENTER); // 垂直置中
        // 設定框線 
        styleRow1.setBorderBottom((short) 1);
        styleRow1.setBorderTop((short) 1);
        styleRow1.setBorderLeft((short) 1);
        styleRow1.setBorderRight((short) 1);
        styleRow1.setWrapText(true); // 自動換行
        /* Title */
        Sheet sheet = workbook.createSheet("檢定名冊");
        sheet.autoSizeColumn(0); // 自動調整欄位寬度
//        sheet.setColumnWidth(0, CHAR_SIZE * Constants.TEN);
//        sheet.setColumnWidth(Constants.ONE, CHAR_SIZE * Constants.TEN);
//        sheet.setColumnWidth(Constants.TWO, CHAR_SIZE * Constants.FIFTEEN);

        Row rowTitle = sheet.createRow(0);
        rowTitle.createCell(0).setCellValue("編號");
        rowTitle.createCell(1).setCellValue("姓名");
        rowTitle.createCell(2).setCellValue("身分證字號");

//        for (int i = 0; i < examineeList.size(); i++) {
            Row rowContent = sheet.createRow(1); // 建立儲存格
            Cell cellContent = rowContent.createCell(0);
            cellContent.setCellValue(1);
            cellContent = rowContent.createCell(1);
            cellContent.setCellValue("john");
            cellContent = rowContent.createCell(2);
            cellContent.setCellValue("H111111");
            //        }
               ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
               workbook.write(outByteStream);
               byte [] outArray = outByteStream.toByteArray();
               response.setContentType("application/ms-excel");
               response.setContentLength(outArray.length);
               response.setHeader("Expires:", "0"); // eliminates browser caching
               response.setHeader("Content-Disposition", "attachment; filename=testxls.xlsx");
               OutputStream outStream = response.getOutputStream();
               outStream.write(outArray);
               outStream.flush();

    }
}
