package acmr.excel.pojo;

import java.io.Serializable;

/**
 * excel边框 ，类型和颜色
 * 
 * @author zengqu
 * 
 */
public class Excelborder implements Cloneable, Serializable {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private short sort; // 线类型，0为无线，1表示实线
	private ExcelColor color; // 线的颜色值

	/**
	 * 构造函数，默认为没有边框，黑色
	 */
	public Excelborder() {
		sort = 0;
		color =null;
	}

	@Override
	public Excelborder clone() {
		Excelborder o = new Excelborder();
		o.color = this.color;
		o.sort = sort;
		return o;

	}

	/**
	 * 返回边框类型
	 * 
	 * @return 边框类型
	 */
	public short getSort() {
		return sort;
	}

	/**
	 * 设置边框类型 ,主要关注是有还是无就好
	 * 
	 * @param sort
	 */
	public void setSort(short sort) {
		this.sort = sort;
	}

	/**
	 * 返回边框颜色
	 * 
	 * @return
	 */
	public ExcelColor getColor() {
		return color;
	}

	/**
	 * 设置边框颜色
	 * 
	 * @param color
	 */
	public void setColor(ExcelColor color) {
		this.color = color;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		}
		if (this.getClass() != obj.getClass()) {
			return false;
		}
		Excelborder o = (Excelborder) obj;
		if (this.color != null) {
			if (this.sort != o.sort || !this.color.equals(o.color)) {
				return false;
			}
		}else{
			if(this.sort!=o.sort || o.color!=null){
				return false;
			}
		}
		return true;
	}

	@Override
	public String toString() {
		return "" + sort + " " + color;
	}

}
