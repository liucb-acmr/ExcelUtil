package acmr.excel.pojo;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import acmr.excel.ExcelException;
import acmr.util.IKeyible;

/**
 * excel行
 * 
 * @author zengqu
 * 
 */
public class ExcelRow implements IKeyible, Cloneable, Serializable {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private int height; // 网页上行高 poi高度=网页高度*18 行高 实际高度*20

	private String code; // 唯一编号

	private List<ExcelCell> cells; // 行中的单元格集合

	private Map<String, String> exps; // 扩展属性

	private boolean rowhidden;

	private boolean inlist; // 是否已经放到队列中了，如果放在队列中了就不能再修改code
	
	private ExcelCellStyle cellstyle;// 行样式


	public ExcelRow() {
		cells = new ArrayList<ExcelCell>();
		exps = new HashMap<String, String>();
		code = "-1";
		// height = 270;
		height = 19;
		inlist = false;
		rowhidden = false;
		cellstyle = new ExcelCellStyle();
	}

	/**
	 * 构造函数，默认高270 为13.5
	 */
	public ExcelRow(String code1) {
		cells = new ArrayList<ExcelCell>();
		exps = new HashMap<String, String>();
		code = code1;
		// height = 270;
		height = 19;
		inlist = false;
		rowhidden = false;
		cellstyle = new ExcelCellStyle();
	}

	@Override
	public ExcelRow clone() {
		ExcelRow o = new ExcelRow(this.code);
		o.height = this.height;
		o.code = this.code;
		o.inlist = false;
		o.rowhidden = this.rowhidden;
		for (int i = 0; i < cells.size(); i++) {
			if (cells.get(i) != null) {
				o.cells.add(cells.get(i).clone());
			} else {
				o.cells.add(null);
			}
		}
		for (String key : this.exps.keySet()) {
			o.exps.put(key, this.exps.get(key));
		}
		return o;
	}

	/**
	 * 扩展属性
	 * 
	 * @return
	 */
	public Map<String, String> getExps() {
		return exps;
	}

	public void setCells(List<ExcelCell> cells) {
		this.cells = cells;
	}

	public void setExps(Map<String, String> exps) {
		this.exps = exps;
	}

	public boolean isRowhidden() {
		return rowhidden;
	}

	public void setRowhidden(boolean rowhidden) {
		this.rowhidden = rowhidden;
	}

	/**
	 * 行高 实际高度*20
	 * 
	 * @return
	 */
	public int getHeight() {
		return height;
	}

	/**
	 * 行高 实际高度*20
	 * 
	 * @param height
	 */
	public void setHeight(int height) {
		this.height = height;
	}

	public String getCode() {
		return code;
	}

	/**
	 * 设置行的唯一编号，请一定在加入sheet前设置
	 * 
	 * @param code
	 * @throws ExcelException
	 */
	public void setCode(String code) throws ExcelException {
		if (this.inlist) {
			throw (new ExcelException("已经不能修改编码值了！"));
		}
		this.code = code;
	}

	/**
	 * 行中的单元格集合
	 * 
	 * @return
	 */
	public List<ExcelCell> getCells() {
		return cells;
	}

	/**
	 * 增加单元格
	 * 
	 * @param e
	 */
	public void add(ExcelCell e) {
		this.cells.add(e);
	}

	/**
	 * 单元格付新值
	 * 
	 * @param col
	 * @param e
	 */
	public void set(int col, ExcelCell e) {
		this.cells.set(col, e);
	}

	/**
	 * 在固定位置新增
	 * 
	 * @param col
	 * @param e
	 */
	public void add(int col, ExcelCell e) {
		this.cells.add(col, e);
	}

	@Override
	public String toString() {
		return cells.toString();
	}

	@Override
	public String Key() {
		return code;
	}

	protected boolean isInlist() {
		return inlist;
	}

	protected void setInlist(boolean inlist) {
		this.inlist = inlist;
	}

	public boolean isNull() {
		for (int i = 0; i < this.cells.size(); i++) {
			if (this.cells.get(i) != null) {
				return false;
			}
		}
		return true;
	}

	public ExcelCellStyle getCellstyle() {
		return cellstyle;
	}

	public void setCellstyle(ExcelCellStyle cellstyle) {
		this.cellstyle = cellstyle;
	}
	
}
