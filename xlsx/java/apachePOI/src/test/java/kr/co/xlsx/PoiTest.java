package kr.co.xlsx;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import kr.co.xlsx.util.CustomerExcelReader;
import kr.co.xlsx.util.CustomerExcelWriter;
import kr.co.xlsx.vo.CustomerVo;


public class PoiTest {
	
	private String s = File.separator; //디렉토리 구분자
	private String route = 
			"C:" +s+ "kyh" +s+ "export" +s+ "excel" +s;
	private String xlsFile = "test.xls";
	private String xlsxFile = "test.xlsx";
	
	@Test
	public void writeExcel() {
		// 엑셀로 쓸 데이터 생성
		List<CustomerVo> list = new ArrayList<CustomerVo>();
		list.add(new CustomerVo("user01", "사용자1", "20", "10yeony@gmail.com"));
		list.add(new CustomerVo("user02", "사용자2", "61", "10yeony@gmail.com"));
		list.add(new CustomerVo("user03", "사용자3", "78", "10yeony@gmail.com"));
		list.add(new CustomerVo("user04", "사용자4", "51", "10yeony@gmail.com"));
		list.add(new CustomerVo("user05", "사용자5", "17", "10yeony@gmail.com"));
		
		CustomerExcelWriter excelWriter = new CustomerExcelWriter();
		//xls 파일 쓰기
		excelWriter.xlsWiter(list);
		
		//xlsx 파일 쓰기
		excelWriter.xlsxWiter(list);
	}
	
	@Test
	public void readExcel() {
		CustomerExcelReader excelReader = new CustomerExcelReader();
		
		System.out.println("*****xls*****");
		List<CustomerVo> xlsList = excelReader.xlsToCustomerVoList(route + xlsFile);
		printList(xlsList );
		
		System.out.println("*****xlsx*****");
		List<CustomerVo> xlsxList = excelReader.xlsxToCustomerVoList(route + xlsxFile);
		printList(xlsxList );
	}
	
	public static void printList(List<CustomerVo> list) {
		CustomerVo vo;
		for (int i = 0; i < list.size(); i++) {
			vo = list.get(i);
			System.out.println(vo.toString());
		}
	}
	
}
