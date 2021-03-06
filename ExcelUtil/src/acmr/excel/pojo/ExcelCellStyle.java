package acmr.excel.pojo;

import java.io.Serializable;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import acmr.excel.ExcelHelper;

/**
 * excel单元格样式
 * 
 * @author zengqu
 * 
 */
public class ExcelCellStyle implements Cloneable, Serializable {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private short align; // 水平方式 1 2 3
	private short valign; // 垂直方式 0 1 2
	private Excelborder topborder; // 上边框
	private Excelborder leftborder; // 左边框
	private Excelborder rightborder;// 右边框
	private Excelborder bottomborder;// 下边框
	private ExcelFont font;// 字体
	private ExcelColor bgcolor;// 背景色
	private ExcelColor fgcolor;// 前景色 一般是这个表示填充色
	private short pattern; // 填充模式
	private String dataformat; // 格式化字符
	private boolean hidden; // 单元格是否隐藏
	private short indention; // 首行缩进
	private short rotation; // 文字的方向
	private boolean shrinktofit; // 是否自动压缩以适用单元格
	private boolean wraptext; // 是否自动换行
	private boolean locked; //是否锁定

	/**
	 * 构造函数，默认
	 */
	public ExcelCellStyle() {
		align = 0;
		bgcolor = null;
		fgcolor = null;

		valign = 1;
		font = new ExcelFont();
		pattern = 0;
		dataformat = "General";
		hidden = false;
		indention = 0;
		shrinktofit = false;
		wraptext = false;
		locked = true;
		this.topborder = new Excelborder();
		this.leftborder = new Excelborder();
		this.rightborder = new Excelborder();
		this.bottomborder = new Excelborder();
	}

	/**
	 * 构造函数，从XSSFCellStyle转换来
	 * 
	 * @param c1
	 * @see XSSFCellStyle
	 */
	public ExcelCellStyle(XSSFCellStyle c1) {
		this.align = c1.getAlignment();
		this.valign = c1.getVerticalAlignment();
		this.topborder = ExcelHelper.getExcelBorder(c1.getBorderTop(), c1.getTopBorderXSSFColor());
		this.leftborder = ExcelHelper.getExcelBorder(c1.getBorderLeft(), c1.getLeftBorderXSSFColor());
		this.rightborder = ExcelHelper.getExcelBorder(c1.getBorderRight(), c1.getRightBorderXSSFColor());
		this.bottomborder = ExcelHelper.getExcelBorder(c1.getBorderBottom(), c1.getBottomBorderXSSFColor());
		this.dataformat = c1.getDataFormatString();
		c1.getDataFormat();
		if (this.dataformat == null) {
			dataformat = "General";
		}
		this.bgcolor = ExcelHelper.getJavaColor(c1.getFillBackgroundXSSFColor(),false);
		this.fgcolor = ExcelHelper.getJavaColor(c1.getFillForegroundXSSFColor(),false);
		this.pattern = c1.getFillPattern();
		if(pattern>0 && fgcolor==null){
			fgcolor=new ExcelColor(255,255,255);
		}
		this.font = new ExcelFont(c1.getFont());
		this.hidden = c1.getHidden();
		this.indention = c1.getIndention();
		this.rotation = c1.getRotation();
		this.shrinktofit = c1.getShrinkToFit();
		this.wraptext = c1.getWrapText();
		this.locked = c1.getLocked();
	}

	@Override
	public ExcelCellStyle clone() {
		ExcelCellStyle o = new ExcelCellStyle();
		o.align = this.align;
		if (this.bgcolor != null) {
			o.bgcolor = this.bgcolor.clone();
		}
		o.bottomborder = this.bottomborder.clone();
		o.dataformat = this.dataformat;
		if (this.fgcolor != null) {
			o.fgcolor = this.fgcolor.clone();
		}
		o.font = this.font.clone();
		o.hidden = this.hidden;
		o.indention = this.indention;
		o.leftborder = this.leftborder.clone();
		o.pattern = this.pattern;
		o.rightborder = this.rightborder.clone();
		o.rotation = this.rotation;
		o.shrinktofit = this.shrinktofit;
		o.topborder = this.topborder.clone();
		o.valign = this.valign;
		o.wraptext = this.wraptext;
		o.locked = this.locked;
		return o;
	}

	/**
	 * 把样式传导到excel中
	 * 
	 * @param xs1
	 * @param book
	 */
	public void setXSSFCellStyle(XSSFCellStyle xs1, XSSFWorkbook book, List<ExcelFont> fonts) {
		xs1.setAlignment(this.align);
		xs1.setBorderTop(this.getTopborder().getSort());
		if (this.getTopborder().getColor() != null) {
			xs1.setTopBorderColor(ExcelHelper.getExcelColor(this.getTopborder().getColor(), false));
		}
		xs1.setBorderLeft(this.getLeftborder().getSort());
		if (this.getLeftborder().getColor() != null) {
			xs1.setLeftBorderColor(ExcelHelper.getExcelColor(this.getLeftborder().getColor(), false));
		}
		xs1.setBorderRight(this.getRightborder().getSort());
		if (this.getRightborder().getColor() != null) {
			xs1.setRightBorderColor(ExcelHelper.getExcelColor(this.getRightborder().getColor(), false));
		}
		xs1.setBorderBottom(this.getBottomborder().getSort());
		if (this.getBottomborder().getColor() != null) {
			xs1.setBottomBorderColor(ExcelHelper.getExcelColor(this.getBottomborder().getColor(), false));
		}
		xs1.setDataFormat(book.createDataFormat().getFormat(this.dataformat));
		xs1.setVerticalAlignment(valign);
		if (bgcolor != null) {
			xs1.setFillBackgroundColor(ExcelHelper.getExcelColor(bgcolor, false));
		}
		if (fgcolor != null) {
			xs1.setFillForegroundColor(ExcelHelper.getExcelColor(fgcolor, false));
		}
		xs1.setFillPattern(pattern);
		xs1.setHidden(hidden);
		xs1.setIndention(indention);
		xs1.setRotation(rotation);
		xs1.setWrapText(wraptext);
		xs1.setLocked(locked);
		int pos = this.findFont(fonts, this.font);
		XSSFFont font1 = null;
		if (pos < 0) {
			font1 = book.createFont();
			this.font.setXSSFFont(font1);
			fonts.add(this.font);
		} else {
			font1 = book.getFontAt((short) (pos + 1));
		}
		xs1.setFont(font1);
	}

	private int findFont(List<ExcelFont> fonts, ExcelFont font1) {
		int pos = -1;
		for (int i = 0; i < fonts.size(); i++) {
			if (fonts.get(i).equals(font1)) {
				pos = i;
				break;
			}
		}
		return pos;
	}

	/**
	 * 水平方式 0 1 2
	 * 
	 * @return
	 */
	public short getAlign() {
		return align;
	}

	/**
	 * 水平方式 0 1 2
	 * 
	 * @param align
	 */
	public void setAlign(short align) {
		this.align = align;
	}

	/**
	 * 垂直方式 0 1 2
	 * 
	 * @return
	 */
	public short getValign() {
		return valign;
	}

	/**
	 * 垂直方式 0 1 2
	 * 
	 * @param valign
	 */
	public void setValign(short valign) {
		this.valign = valign;
	}

	/**
	 * 上边框
	 * 
	 * @return
	 */
	public Excelborder getTopborder() {
		return topborder;
	}

	/**
	 * 上边框
	 * 
	 * @param topborder
	 */
	public void setTopborder(Excelborder topborder) {
		this.topborder = topborder;
	}

	/**
	 * 左边框
	 * 
	 * @return
	 */
	public Excelborder getLeftborder() {
		return leftborder;
	}

	/**
	 * 左边框
	 * 
	 * @param leftborder
	 */
	public void setLeftborder(Excelborder leftborder) {
		this.leftborder = leftborder;
	}

	/**
	 * 右边框
	 * 
	 * @return
	 */
	public Excelborder getRightborder() {
		return rightborder;
	}

	/**
	 * 右边框
	 * 
	 * @param rightborder
	 */
	public void setRightborder(Excelborder rightborder) {
		this.rightborder = rightborder;
	}

	/**
	 * 下边框
	 * 
	 * @return
	 */
	public Excelborder getBottomborder() {
		return bottomborder;
	}

	/**
	 * 下边框
	 * 
	 * @param bottomborder
	 */
	public void setBottomborder(Excelborder bottomborder) {
		this.bottomborder = bottomborder;
	}

	/**
	 * 字体
	 * 
	 * @return
	 */
	public ExcelFont getFont() {
		return font;
	}

	/**
	 * 字体
	 * 
	 * @param font
	 */
	public void setFont(ExcelFont font) {
		this.font = font;
	}

	/**
	 * 背景色
	 * 
	 * @return
	 */
	public ExcelColor getBgcolor() {
		return bgcolor;
	}

	/**
	 * 背景色
	 * 
	 * @param bgcolor
	 */
	public void setBgcolor(ExcelColor bgcolor) {
		this.bgcolor = bgcolor;
	}

	/**
	 * 前景色 一般是这个表示填充色
	 * 
	 * @return
	 */
	public ExcelColor getFgcolor() {
		return fgcolor;
	}

	/**
	 * 前景色 一般是这个表示填充色
	 * 
	 * @param fgcolor
	 */
	public void setFgcolor(ExcelColor fgcolor) {
		this.fgcolor = fgcolor;
	}

	/**
	 * 填充模式
	 * 
	 * @return
	 */
	public short getPattern() {
		return pattern;
	}

	/**
	 * 填充模式
	 * 
	 * @param pattern
	 */
	public void setPattern(short pattern) {
		this.pattern = pattern;
	}

	/**
	 * 格式化字符
	 * 
	 * @return
	 */
	public String getDataformat() {
		return dataformat;
	}

	/**
	 * 格式化字符
	 * 
	 * @param dataformat
	 */
	public void setDataformat(String dataformat) {
		this.dataformat = dataformat;
	}

	/**
	 * 单元格是否隐藏
	 * 
	 * @return
	 */
	public boolean isHidden() {
		return hidden;
	}

	/**
	 * 单元格是否隐藏
	 * 
	 * @param hidden
	 */
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	/**
	 * 首行缩进
	 * 
	 * @return
	 */
	public short getIndention() {
		return indention;
	}

	/**
	 * 首行缩进
	 * 
	 * @param indention
	 */
	public void setIndention(short indention) {
		this.indention = indention;
	}

	/**
	 * 文字的方向
	 * 
	 * @return
	 */
	public short getRotation() {
		return rotation;
	}

	/**
	 * 文字的方向
	 * 
	 * @param rotation
	 */
	public void setRotation(short rotation) {
		this.rotation = rotation;
	}

	/**
	 * 是否自动压缩以适用单元格
	 * 
	 * @return
	 */
	public boolean isShrinktofit() {
		return shrinktofit;
	}

	/**
	 * 是否自动压缩以适用单元格
	 * 
	 * @param shrinktofit
	 */
	public void setShrinktofit(boolean shrinktofit) {
		this.shrinktofit = shrinktofit;
	}

	/**
	 * 是否自动换行
	 * 
	 * @return
	 */
	public boolean isWraptext() {
		return wraptext;
	}

	/**
	 * 是否自动换行
	 * 
	 * @param wraptext
	 */
	public void setWraptext(boolean wraptext) {
		this.wraptext = wraptext;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		}
		if (this.getClass() != obj.getClass()) {
			return false;
		}
		ExcelCellStyle o = (ExcelCellStyle) obj;
		if (this.hidden != o.hidden || this.shrinktofit != o.shrinktofit || this.wraptext != o.wraptext || this.locked != o.locked) {
			return false;
		}
		if (this.align != o.align || this.indention != o.indention || this.pattern != o.pattern || this.rotation != o.rotation || this.valign != o.valign) {
			return false;
		}
		if (this.bgcolor == null && o.bgcolor != null) {
			return false;
		}
		if (this.fgcolor == null && o.fgcolor != null) {
			return false;
		}
		if ((this.bgcolor != null && !this.bgcolor.equals(o.bgcolor)) || !this.dataformat.equals(o.dataformat) || (this.fgcolor != null && !this.fgcolor.equals(o.fgcolor))) {
			return false;
		}
		if (!this.topborder.equals(o.topborder) || !this.leftborder.equals(o.leftborder) || !this.rightborder.equals(o.rightborder) || !this.bottomborder.equals(o.bottomborder)) {
			return false;
		}
		if (!this.font.equals(o.font)) {
			return false;
		}
		return true;
	}

	public boolean isLocked() {
		return locked;
	}

	public void setLocked(boolean locked) {
		this.locked = locked;
	}
}
