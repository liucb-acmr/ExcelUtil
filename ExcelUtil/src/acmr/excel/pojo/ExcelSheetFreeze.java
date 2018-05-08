package acmr.excel.pojo;

import java.io.Serializable;

public class ExcelSheetFreeze  implements Cloneable,Serializable{
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private boolean freeze;   //true为冻结，false 为分隔
	private int row;
	private int col;
	private int firstrow;
	private int firstcol;
	
	private int activepan;

	public ExcelSheetFreeze() {
		freeze = true;
		row = 0;
		col = 0;
		firstrow = 0;
		firstcol = 0;
	}

	@Override
	public ExcelSheetFreeze clone() {
		ExcelSheetFreeze o =new ExcelSheetFreeze();
		o.freeze=this.freeze;
		o.row=this.row;
		o.col=this.col;
		o.firstrow=this.firstrow;
		o.firstcol=this.firstcol;
		o.activepan=this.activepan;
		return o;
		
	}
	public boolean isFreeze() {
		return freeze;
	}

	public void setFreeze(boolean freeze) {
		this.freeze = freeze;
	}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getCol() {
		return col;
	}

	public void setCol(int col) {
		this.col = col;
	}

	public int getFirstrow() {
		return firstrow;
	}

	public void setFirstrow(int firstrow) {
		this.firstrow = firstrow;
	}

	public int getFirstcol() {
		return firstcol;
	}

	public void setFirstcol(int firstcol) {
		this.firstcol = firstcol;
	}

	public int getActivepan() {
		return activepan;
	}

	public void setActivepan(int activepan) {
		this.activepan = activepan;
	}

}
