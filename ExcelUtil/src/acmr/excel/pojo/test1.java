package acmr.excel.pojo;

import java.text.DecimalFormat;

import org.apache.poi.xssf.usermodel.XSSFColor;

import acmr.util.PubInfo;

public class test1 {
	int int1;
	String str1;

	public int getInt1() {
		return int1;
	}

	public void setInt1(int int1) {
		this.int1 = int1;
	}

	public String getStr1() {
		return str1;
	}

	public void setStr1(String str1) {
		this.str1 = str1;
	}

	 

	public static void main(String[] args) {
		XSSFColor xc1 = new XSSFColor();
		byte[] bs = new byte[]{(byte) 255,(byte) 255,(byte) 255};
		xc1.setRgb(bs);
		PubInfo.printStr(xc1.getRgb().toString());
		
		
	}

	
	
}
