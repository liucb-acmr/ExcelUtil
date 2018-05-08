package acmr.excel.pojo;

import java.io.Serializable;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import acmr.excel.pojo.Constants.CELLTYPE;

/**
 * excel单元格
 * 
 * @author zengqu
 * 
 */
/**
 * @author zengqu
 *
 */
/**
 * @author zengqu
 * 
 */
public class ExcelCell implements Cloneable, Serializable {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private int rowspan; // 跨行数
	private int colspan; // 跨列数

	private String text; // 单元格内容

	private Object value; // 值 注意日期的问题，excel的开始时间为1900-1-0 而java为1900-1-1
 
	private CELLTYPE type; // 单元格文本类型 数字，文本，时间，等等

	private String memo; // 单元格的备注

	private Map<String, String> exps; // 扩展属性

	private ExcelCellStyle cellstyle;// 单元格样式

	/**
	 * 构造函数，默认一个空单元格
	 */
	public ExcelCell() {
		rowspan = 1;
		colspan = 1;
		text = "";
		exps = new HashMap<String, String>();
		memo = "";
		cellstyle = new ExcelCellStyle();
		type = CELLTYPE.BLANK;
	}

	@Override
	public ExcelCell clone() {
		ExcelCell o = new ExcelCell();
		o.rowspan = rowspan;
		o.colspan = this.colspan;
		o.text = this.text;
		o.memo = this.memo;
		o.type = this.type;
		o.value = this.value;
	 
		if (this.cellstyle != null) {
			o.cellstyle = this.cellstyle.clone();
		}
		for (String key : this.exps.keySet()) {
			o.exps.put(key, this.exps.get(key));
		}
		return o;
	}

	public String getShowText() {
		return ExcelFormat.getShowText(this);
	}
	/**
	 * 返回行的跨度，1 表示不跨行
	 * 
	 * @return 行的跨度
	 */
	public int getRowspan() {
		return rowspan;
	}

	/**
	 * 设置行的跨度
	 * 
	 * @param rowspan
	 */
	public void setRowspan(int rowspan) {
		if (rowspan < 1) {
			rowspan = 1;
		}
		this.rowspan = rowspan;
	}

	/**
	 * 返回列的跨度
	 * 
	 * @return
	 */
	public int getColspan() {
		return colspan;
	}

	/**
	 * 设置列的跨度
	 * 
	 * @param colspan
	 */
	public void setColspan(int colspan) {
		if (colspan < 1) {
			colspan = 1;
		}
		this.colspan = colspan;
	}

	/**
	 * 返回文本，单元格的原始文本
	 * 
	 * @return 原始文本
	 */
	public String getText() {
		return text;
	}

	/**
	 * 设置文本
	 * 
	 * @param text
	 */
	public void setText(String text) {
		this.text = text;
	}

	/**
	 * 返回值
	 * 
	 * @return
	 */
	public Object getValue() {
		return value;
	}

	/**
	 * 设置值
	 * 
	 * @param value
	 */
	public void setValue(Object value) {
		this.value = value;
	}
	
	public void setCellValue(Object value){
		this.cellstyle.setDataformat("General");
		if (value instanceof Date) {
			this.type = CELLTYPE.DATE;
			this.cellstyle.setDataformat("yyyy/m/d");
		} else if (value instanceof Double) {
			this.type = CELLTYPE.NUMERIC;
		} else if (value instanceof Long || value instanceof Integer) {
			value = Double.parseDouble(value.toString());
			this.type = CELLTYPE.NUMERIC;
		} else {
			this.type = CELLTYPE.STRING;
		}
		this.value = value;
		this.text=value.toString();
	}

	/**
	 * 单元格值类型
	 * 
	 * @return type
	 * @see CELLTYPE
	 */
	public CELLTYPE getType() {
		return type;
	}

	/**
	 * 单元格类型
	 * 
	 * @param type
	 */
	public void setType(CELLTYPE type) {
		this.type = type;
	}

	/**
	 * 单元格备注
	 * 
	 * @return
	 */
	public String getMemo() {
		return memo;
	}

	/**
	 * 单元格备注
	 * 
	 * @param memo
	 */
	public void setMemo(String memo) {
		this.memo = memo;
	}

	/**
	 * 单元格样式
	 * 
	 * @return
	 * @see ExcelCellStyle
	 */
	public ExcelCellStyle getCellstyle() {
		return cellstyle;
	}

	/**
	 * 单元格样式
	 * 
	 * @param cellstyle
	 */
	public void setCellstyle(ExcelCellStyle cellstyle) {
		this.cellstyle = cellstyle;
	}

	/**
	 * 扩展属性
	 * 
	 * @return
	 */
	public Map<String, String> getExps() {
		return exps;
	}

	public void setExps(Map<String, String> exps) {
		this.exps = exps;
	}

	@Override
	public String toString() {
		return text;
	}

}
