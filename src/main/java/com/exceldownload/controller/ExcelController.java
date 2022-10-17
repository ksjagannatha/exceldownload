package com.exceldownload.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import com.exceldownload.service.ExcelService;

@RestController
public class ExcelController {

	@Autowired private ExcelService excelService;
	
	@PostMapping("/download")
	public ResponseEntity<?> writeExcel(@RequestBody String requestBody){
		return new ResponseEntity<>(excelService.getDownload(requestBody), HttpStatus.OK);
	}
}
