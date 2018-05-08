package acmr.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelSheet;

public class Test {

	public static void main(String[] args) throws IOException {
		FileInputStream inp = null;
		try {
			inp = new FileInputStream("D:\\a.xlsx");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		XSSFWorkbook wb = new XSSFWorkbook(inp);
		Sheet sheet = wb.getSheetAt(0); // 获得第三个工作薄(2008工作薄)
		int list = sheet.getNumMergedRegions();
		CellRangeAddress ca = sheet.getMergedRegion(0);
		Row row = sheet.getRow(3);
		XSSFCell cell =  (XSSFCell) row.getCell(2);
		
		
		ExcelSheet es = new ExcelSheet();
  	    ExcelCell ec = es.getExcelCell(cell);
		
		System.out.println(list);
		// 填充上面的表格,数据需要从数据库查询
	/*	XSSFRow row5 = sheet.getRow(4); // 获得工作薄的第五行
		XSSFCell cell54 = row5.getCell(3);// 获得第五行的第四个单元格
		cell54.setCellValue("测试纳税人名称");// 给单元格赋值
		//获得总列数
		int coloumNum=sheet.getRow(0).getPhysicalNumberOfCells();
		int rowNum=sheet.getLastRowNum();//获得总行数
*/
		
		/*List<Object> list =  new ArrayList<Object>();
		list.add("a");
		list.add("b");
		ExcelColor  ec = new ExcelColor();
		ec.setB(1);
		list.add(ec);
		ExcelColor  ec1 = new ExcelColor();
		ec1.setB(1);
		
		System.out.println(list.indexOf(ec1.getB()));*/
	}

}
