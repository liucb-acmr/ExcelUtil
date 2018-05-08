package acmr.excel.pojo;

import java.io.Serializable;

public class ExcelCellRangeAddress implements Serializable{
	private int FirstColumn;
	private int FirstRow;
	private int LastColumn;
	private int LastRow;

	public int getFirstColumn() {
		return FirstColumn;
	}

	public void setFirstColumn(int firstColumn) {
		FirstColumn = firstColumn;
	}

	public int getFirstRow() {
		return FirstRow;
	}

	public void setFirstRow(int firstRow) {
		FirstRow = firstRow;
	}

	public int getLastColumn() {
		return LastColumn;
	}

	public void setLastColumn(int lastColumn) {
		LastColumn = lastColumn;
	}

	public int getLastRow() {
		return LastRow;
	}

	public void setLastRow(int lastRow) {
		LastRow = lastRow;
	}

}
