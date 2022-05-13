package com.suresh;

import com.suresh.service.AsposeService;
import com.suresh.service.DemoService;
import com.suresh.service.FormatRow;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class TestProjApplication {

	public static void main(String[] args) throws Exception {
		SpringApplication.run(TestProjApplication.class, args);
//		AsposeService asposeService = new AsposeService();
//		asposeService.setStyles();
//		FormatRow formatRow = new FormatRow();
//		formatRow.formatRow();
		DemoService demoService = new DemoService();
		//demoService.numberToText();
		demoService.textToNumber();
	}

}
