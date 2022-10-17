package com.exceldownload.serviceimpl;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import com.exceldownload.constant.ConstantData;
import com.exceldownload.service.ExcelService;

@Service
public class ExcelServiceImpl implements ExcelService {

	
	private static final Logger logger=LoggerFactory.getLogger(ExcelServiceImpl.class);
	
	@Value("${filepa}")
	private String filepath;
	
	@Override
	public Map<String, String> getDownload(String requestBody) {
		Map<String, String> map=new HashMap<>();
		FileOutputStream outPutStream=null;
		XSSFWorkbook workBook=new XSSFWorkbook();
		logger.info(ConstantData.INITIALIZE);
		try {
			XSSFSheet sheet=workBook.createSheet(ConstantData.SHEET_NAME);
			Font headerFont=workBook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short)12);
			headerFont.setColor(IndexedColors.BLACK.index);
			CellStyle headerStyle=workBook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			JSONArray array=new JSONArray(requestBody);
			JSONObject obj= array.getJSONObject(0);
			List<String> list= new ArrayList<>(obj.keySet());
			Row headerRow=sheet.createRow(0);
			for(int i=0;i<list.size();i++) {
				Cell cell=headerRow.createCell(i);
				cell.setCellValue(list.get(i));
				cell.setCellStyle(headerStyle);
			}
			
			for(int i=0;i<array.length();i++) {
				JSONObject objec=array.getJSONObject(i);
				Row header=sheet.createRow(i+1);
				Iterator<String> para= objec.keySet().iterator();
				for(int j=0;j<obj.length();j++) {
					String key=para.next();
					Cell cell=header.createCell(j);
					cell.setCellValue(objec.get(key).toString());
				}
			}
			
			for(int i=0;i<list.size();i++) {
				sheet.autoSizeColumn(i);
			}
			outPutStream=new FileOutputStream(filepath);
			workBook.write(outPutStream);
			
			map.put(ConstantData.MESSAGE, ConstantData.SUCCESS);
			
		} catch (Exception e) {
			logger.error(ConstantData.ERROR,e);
			
			map.put(ConstantData.MESSAGE, ConstantData.FAILED);
		}
		finally {
			try {
				if(outPutStream!=null) {
				outPutStream.close();
				workBook.close();
				}
			} catch (IOException e) {
				
				logger.error(ConstantData.CLOSE,e);
				
			}
			
		}
		return map;
	}

}
