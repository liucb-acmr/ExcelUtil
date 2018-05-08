package acmr.excel.pojo;

import java.io.Serializable;

import org.apache.poi.xssf.usermodel.XSSFFont;

import acmr.excel.ExcelHelper;

/**
 * excel字体
 * 
 * @author zengqu
 * 
 */
public class ExcelFont implements Cloneable, Serializable {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private String fontname; // 字体名字
	private short size;// 字体字号 实际字号*20
	private short boldweight; // 字体黑色浓度
	private ExcelColor color; // 字体颜色
	private boolean strikeout; // 删除线
	private byte underline; // 下划线
	private boolean italic; // 斜体
	private short typeoffset; // 角标类型

	/**
	 * 构造函数
	 */
	public ExcelFont() {
		fontname = "宋体";
		size = 220;
		boldweight = 400;
		color = null;
		strikeout = false;
		underline = 0;
		italic = false;
		typeoffset = 0;
	}

	/**
	 * 构造函数，从poi取得字体
	 * 
	 * @param f1
	 */
	public ExcelFont(XSSFFont f1) {
		this.boldweight = f1.getBoldweight();
		this.color = ExcelHelper.getJavaColor(f1.getXSSFColor(),true);
		this.fontname = f1.getFontName();
		this.italic = f1.getItalic();
		this.size = f1.getFontHeight();
		this.strikeout = f1.getStrikeout();
		this.typeoffset = f1.getTypeOffset();
		this.underline = f1.getUnderline();
	}

	@Override
	public ExcelFont clone() {
		ExcelFont o = new ExcelFont();
		o.boldweight = this.boldweight;
		o.color = this.color;
		o.fontname = this.fontname;
		o.italic = this.italic;
		o.size = this.size;
		o.strikeout = this.strikeout;
		o.typeoffset = this.typeoffset;
		o.underline = this.underline;
		return o;
	}

	/**
	 * 把字体信息写入poi对象
	 * 
	 * @param f1
	 */
	public void setXSSFFont(XSSFFont f1) {
		f1.setBoldweight(boldweight);
		if (color != null) {
			f1.setColor(ExcelHelper.getExcelColor(color, true));
		}
		f1.setFontName(fontname);
		f1.setItalic(italic);
		f1.setFontHeight(size);
		f1.setStrikeout(strikeout);
		f1.setTypeOffset(typeoffset);
		f1.setUnderline(underline);
	}

	/**
	 * 字体名字
	 * 
	 * @return
	 */
	public String getFontname() {
		return fontname;
	}

	/**
	 * 字体名字
	 * 
	 * @param fontname
	 */
	public void setFontname(String fontname) {
		this.fontname = fontname;
	}

	/**
	 * 字体字号 实际字号*20
	 * 
	 * @return
	 */
	public short getSize() {
		return size;
	}

	/**
	 * 字体字号 实际字号*20
	 * 
	 * @param size
	 */
	public void setSize(short size) {
		this.size = size;
	}

	/**
	 * 字体黑色浓度
	 * 
	 * @return
	 */
	public short getBoldweight() {
		return boldweight;
	}

	/**
	 * 字体黑色浓度
	 * 
	 * @param boldweight
	 */
	public void setBoldweight(short boldweight) {
		this.boldweight = boldweight;
	}

	/**
	 * 字体颜色
	 * 
	 * @return
	 */
	public ExcelColor getColor() {
		return color;
	}

	/**
	 * 字体颜色
	 * 
	 * @param color
	 */
	public void setColor(ExcelColor color) {
		this.color = color;
	}

	/**
	 * 删除线
	 * 
	 * @return
	 */
	public boolean isStrikeout() {
		return strikeout;
	}

	/**
	 * 删除线
	 * 
	 * @param strikeout
	 */
	public void setStrikeout(boolean strikeout) {
		this.strikeout = strikeout;
	}

	/**
	 * 下划线
	 * 
	 * @return
	 */
	public byte getUnderline() {
		return underline;
	}

	/**
	 * 下划线
	 * 
	 * @param underline
	 */
	public void setUnderline(byte underline) {
		this.underline = underline;
	}

	/**
	 * 斜体
	 * 
	 * @return
	 */
	public boolean isItalic() {
		return italic;
	}

	/**
	 * 斜体
	 * 
	 * @param italic
	 */
	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	/**
	 * 角标类型
	 * 
	 * @return
	 */
	public short getTypeoffset() {
		return typeoffset;
	}

	/**
	 * 角标类型
	 * 
	 * @param typeoffset
	 */
	public void setTypeoffset(short typeoffset) {
		this.typeoffset = typeoffset;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		}
		if (this.getClass() != obj.getClass()) {
			return false;
		}
		ExcelFont o = (ExcelFont) obj;
		if (this.italic != o.italic || this.strikeout != o.strikeout) {
			return false;
		}
		if (this.boldweight != o.boldweight || this.size != o.size || this.typeoffset != o.typeoffset || this.underline != o.underline) {
			return false;
		}
		if (this.color == null && o.color != null) {
			return false;
		}
		if (this.color != null && !this.color.equals(o.color)) {
			return false;
		}
		if (!this.fontname.equals(o.fontname)) {
			return false;
		}
		return true;
	}

}
