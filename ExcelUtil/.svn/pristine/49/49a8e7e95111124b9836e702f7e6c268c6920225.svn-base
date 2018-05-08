package acmr.excel;

import java.awt.Color;

import org.apache.poi.xssf.usermodel.XSSFColor;

import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.Excelborder;
import acmr.util.PubInfo;

/**
 * excel的一些公用的操作函数
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
public class ExcelHelper {

	/**
	 * 列的编号转换，从索引值到字母表示
	 * 
	 * @param index
	 * @return
	 */
	public static String getColCode(int index) {
		String straz = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		int azlen = straz.length();
		String str1 = "";
		while (index >= 0) {
			int int1 = index - ((index / azlen) * azlen);
			str1 = straz.substring(int1, int1 + 1) + str1;
			index = index / azlen - 1;
		}
		return str1;
	}

	/**
	 * 列的编号转换，从字母表示到索引值
	 * 
	 * @param strbh
	 * @return
	 */
	public static int getColIndex(String strbh) {
		String straz = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		int azlen = straz.length();
		strbh = strbh.toUpperCase();
		int col = -1;
		for (int i = 0; i < strbh.length(); i++) {
			col = col + 1;
			col = col * azlen + straz.indexOf(strbh.substring(i, i + 1));
		}
		return col;
	}

	/**
	 * 返回excel单元格的简写字符 比如B8
	 * 
	 * @param row
	 * @param col
	 * @return
	 */
	public static String getExcelstrBH(int row, int col) {
		String str1 = "";
		str1 = getColCode(col);
		if (row >= 0) {
			str1 = str1 + (row + 1);
		}
		return str1;
	}

	/**
	 * 返回excel对象中的索引号， 索引号是从0开始
	 * 
	 * @param strbh
	 * @return -1 表示没有定义
	 */
	public static int[] getExcelintBH(String strbh) {
		int row = -1;
		int col = -1;
		String strsz = "0123456789";
		int pos = -1;
		for (int i = 0; i < strbh.length(); i++) {
			if (strsz.indexOf(strbh.substring(i, i + 1)) >= 0) {
				pos = i;
				break;
			}
		}
		if (pos < 0) { // 没有行
			col = getColIndex(strbh);
		} else if (pos == 0) { // 没有列
			row = Integer.parseInt(strbh) - 1;
		} else {
			String colstr = strbh.substring(0, pos);
			String rowstr = strbh.substring(pos);
			col = getColIndex(colstr);
			row = Integer.parseInt(rowstr) - 1;
		}
		return new int[] { row, col };
	}

	/**
	 * poi的颜色转成java颜色
	 * 
	 * @param xc
	 * @return
	 */
	public static ExcelColor getJavaColor(XSSFColor xc, boolean isfont) {
		if (xc == null) {
			return null;
		}
		if (xc.isAuto()) {
			return new ExcelColor(0, 0, 0);
		}
		byte[] s = null;
		try {
			s = xc.getRgb();
		} catch (Exception e) {
           e.printStackTrace();
           return null;
		}
		if (s == null) {
			return ExcelColor.getIndexedColor(xc.getIndexed());
		}
		double t = xc.getTint();
		int r = s[0] & 0xff;
		int g = s[1] & 0xff;
		int b = s[2] & 0xff;
		r = getLum(r, t);
		g = getLum(g, t);
		b = getLum(b, t);
		ExcelColor color1 = new ExcelColor(r, g, b);
		return color1;
	}

	public static int getLum(int r, double t) {
		int r1 = 0;
		if (t <= 0) {
			r1 = (int) (r * (1.0 + t));
		} else {
			r1 = (int) (r * (1.0 - t) + (255 - 255 * (1.0 - t)));
		}
		return r1;
	}

	/**
	 * java颜色转成poi颜色，注意字体有特需处理
	 * 
	 * @param c1
	 * @param isfont
	 * @return
	 */
	public static XSSFColor getExcelColor(ExcelColor c1, boolean isfont) {
		if (c1 == null) {
			return null;
		}
		XSSFColor xc1 = new XSSFColor();
		String c2 = c1.getRGBString();
		if (isfont) {
			if (c2.equals("000000")) {
				return null;
			}
		}
		if (c2.equals("000000")) {
			c2 = "FFFFFF";
		} else if (c2.equals("FFFFFF")) {
			c2 = "000000";
		}
		byte[] bs = new byte[3];
		bs[0] = (byte) (0xff & Integer.parseInt(c2.substring(0, 2), 16));
		bs[1] = (byte) (0xff & Integer.parseInt(c2.substring(2, 4), 16));
		bs[2] = (byte) (0xff & Integer.parseInt(c2.substring(4), 16));
		xc1.setRgb(bs);
		xc1.setTint(0);
		return xc1;
	}

	/**
	 * 返回excel单元格的边
	 * 
	 * @param s
	 * @param xc
	 * @return
	 */
	public static Excelborder getExcelBorder(short s, XSSFColor xc) {
		Excelborder line1 = new Excelborder();
		line1.setSort(s);
		s++;
		ExcelColor c1 = ExcelHelper.getJavaColor(xc, false);
		line1.setColor(c1);
		return line1;
	}

	public static void main(String[] args) {
		Color color = new Color(1231231321);
		PubInfo.printStr("" + color.getRGB() + "  " + color.toString());
	}
}
