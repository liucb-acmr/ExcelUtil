package acmr.excel.pojo;

import java.io.Serializable;

public class ExcelColor implements Cloneable, Serializable {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private int r;
	private int g;
	private int b;

	public void setR(int r) {
		this.r = r;
	}

	public void setG(int g) {
		this.g = g;
	}

	public void setB(int b) {
		this.b = b;
	}

	private static final String[] indexedColors = new String[] { "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080",
			"FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333", "000000", "FFFFFF", "646464", "f0f0f0", "000000", "FFFFFF", "A0A0A0",
			"3399FF", "000000", "C8C8C8", "373737", "FFFFFF", "646464", "000000", "FFFFFF", "000000", "FFFFE1", "000000" };

	@Override
	protected ExcelColor clone() {
		ExcelColor o = null;
		try {
			o = (ExcelColor) super.clone();
		} catch (CloneNotSupportedException e) {
			e.printStackTrace();
		}
		return o;
	}

	public ExcelColor() {
		r = 255;
		g = 255;
		b = 255;
	}

	public ExcelColor(int r1, int g1, int b1) {
		r = r1;
		g = g1;
		b = b1;
	}

	public ExcelColor(int[] rgb) {
		r = rgb[0];
		g = rgb[1];
		b = rgb[2];
	}

	public ExcelColor(String rgb) {
		r = (0xff & Integer.parseInt(rgb.substring(0, 2), 16));
		g = (0xff & Integer.parseInt(rgb.substring(2, 4), 16));
		b = (0xff & Integer.parseInt(rgb.substring(4), 16));
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		}
		if (this.getClass() != obj.getClass()) {
			return false;
		}
		ExcelColor o = (ExcelColor) obj;

		if (this.r != o.r || this.g != o.g || this.b != o.b) {
			return false;
		}
		return true;
	}

	public int getR() {
	 
		return r;
	}

	public int getG() {
	 
		return g;
	}

	public int getB() {
		 
		return b;
	}

	public void setRGB(int r, int g, int b) {
		this.b = b;
		this.r = r;
		this.g = g;
 
	}

	public int[] getRGBInt() {
	 
		int[] rgb = new int[3];
		rgb[0] = r;
		rgb[1] = g;
		rgb[2] = b;
		return rgb;
	}

	public String getRGBString() {
	 
		String rgb = "";
		String str1 = "0" + Integer.toHexString(r).toUpperCase();
		rgb += str1.substring(str1.length() - 2);
		str1 = "0" + Integer.toHexString(g).toUpperCase();
		rgb += str1.substring(str1.length() - 2);
		str1 = "0" + Integer.toHexString(b).toUpperCase();
		rgb += str1.substring(str1.length() - 2);
		return rgb;
	}

	public String toString() {
		return   " " + r + " " + g + " " + b;
	}

	public static ExcelColor getIndexedColor(int index) {
		if (indexedColors.length > index && index >= 0) {
			return new ExcelColor(indexedColors[index]);
		}
		return new ExcelColor("FFFFFF");
	}
}
