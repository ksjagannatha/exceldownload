package com.exceldownload;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;

@Configuration
@PropertySource("application.properties")
@SpringBootApplication
public class ExceldownloadApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExceldownloadApplication.class, args);
	}

}
