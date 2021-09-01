package com.example.demo;

import java.util.Scanner;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java_cup.internal_error;
import java_cup.simple_calc.scanner;

@SpringBootApplication
public class ReadExcelApplication {

	public static void main(String[] args) {
		SpringApplication.run(ReadExcelApplication.class, args);
//		ReadExcel readExcel = new ReadExcel();
//		readexcel2 re2= new readexcel2();
		
		try {
//			readExcel.readExcelStart();
//			re2.readExcelStart();
			FindCOMP.test();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
