package kr.co.xlsx;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import kr.co.xlsx.util.CustomerExcelReader;
import kr.co.xlsx.util.CustomerExcelWriter;
import kr.co.xlsx.vo.CustomerVo;


public class PoiTest {
	
//	@Test
//	public void writeExcel() {
//		// 엑셀로 쓸 데이터 생성
//		List<CustomerVo> list = new ArrayList<CustomerVo>();
//		list.add(new CustomerVo("asdf1", "사용자1", "30", "asdf1@naver.com"));
//		list.add(new CustomerVo("asdf2", "사용자2", "31", "asdf2@naver.com"));
//		list.add(new CustomerVo("asdf3", "사용자3", "32", "asdf3@naver.com"));
//		list.add(new CustomerVo("asdf4", "사용자4", "33", "asdf4@naver.com"));
//		list.add(new CustomerVo("asdf5", "사용자5", "34", "asdf5@naver.com"));
//		
//		CustomerExcelWriter excelWriter = new CustomerExcelWriter();
//		//xls 파일 쓰기
//		excelWriter.xlsWiter(list);
//		
//		//xlsx 파일 쓰기
//		excelWriter.xlsxWiter(list);
//	}
	
	@Test
	public void readExcel() {
		CustomerExcelReader excelReader = new CustomerExcelReader();
		
		System.out.println("*****xls*****");
		List<CustomerVo> xlsList = excelReader.xlsToCustomerVoList("");
		printList(xlsList );
		
		System.out.println("*****xlsx*****");
		List<CustomerVo> xlsxList = excelReader.xlsxToCustomerVoList("");
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
