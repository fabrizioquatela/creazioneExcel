package it.cybsec.formazione.creazioneExcel.controller;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import it.cybsec.formazione.creazioneExcel.utility.SharedConstants;

@RestController
@RequestMapping(SharedConstants.EXCEL_CONTROLLER)
public class ExcelController {

	/********************
	 * Excel Settings
	 ********************/
	private final static String EXCEL_SHEET_NAME = "PERSONS";
	
	/********************
	 * API Begin
	 ********************/
	
	@GetMapping(value = SharedConstants.EXCEL_CONTROLLER_CREATE)
	public ResponseEntity<InputStreamResource> create() throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();

		Sheet sheet = workbook.createSheet(EXCEL_SHEET_NAME);
		//sheet.setColumnWidth(0, 6000);
		//sheet.setColumnWidth(1, 4000);

		Row header = sheet.createRow(0);

		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		XSSFFont font = workbook.createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 16);
		font.setBold(true);
		headerStyle.setFont(font);

		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("Nome");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(1);
		headerCell.setCellValue("Cognome");
		headerCell.setCellStyle(headerStyle);
		
		CellStyle style = workbook.createCellStyle();
		style.setWrapText(true);

		Row row = sheet.createRow(2);
		Cell cell = row.createCell(0);
		cell.setCellValue("Mario");
		cell.setCellStyle(style);

		cell = row.createCell(1); 
		cell.setCellValue("Rossi");
		cell.setCellStyle(style);

		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);
		workbook.close();
		
		HttpHeaders headers = new HttpHeaders();
		headers.add("Content-Disposition", "attachment; filename = example.xlsx");
		headers.add("Content-Type", "application/vnd.ms-excel");
		
		ByteArrayInputStream excel = new ByteArrayInputStream(outputStream.toByteArray());
		return ResponseEntity.ok().headers(headers).contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
				.body(new InputStreamResource(excel));
	}
	
}
