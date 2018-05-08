package acmr.excel.pojo;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class ExcelDataValidation implements Serializable{
	private List<ExcelCellRangeAddress> excelCellRangeAddresses = new ArrayList<ExcelCellRangeAddress>();
	private String formula1;
	private String formula2;
	private int operator;
	private int validationType;

	public List<ExcelCellRangeAddress> getExcelCellRangeAddresses() {
		return excelCellRangeAddresses;
	}

	public void setExcelCellRangeAddresses(
			List<ExcelCellRangeAddress> excelCellRangeAddresses) {
		this.excelCellRangeAddresses = excelCellRangeAddresses;
	}

	public String getFormula1() {
		return formula1;
	}

	public void setFormula1(String formula1) {
		this.formula1 = formula1;
	}

	public String getFormula2() {
		return formula2;
	}

	public void setFormula2(String formula2) {
		this.formula2 = formula2;
	}

	public int getOperator() {
		return operator;
	}

	public void setOperator(int operator) {
		this.operator = operator;
	}

	public int getValidationType() {
		return validationType;
	}

	public void setValidationType(int validationType) {
		this.validationType = validationType;
	}

}
