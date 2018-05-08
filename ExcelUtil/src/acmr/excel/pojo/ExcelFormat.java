package acmr.excel.pojo;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Locale;

import org.apache.poi.ss.usermodel.BuiltinFormats;

import acmr.util.PubInfo;

public class ExcelFormat {

	public static String getShowText(ExcelCell cell1) {
		if (cell1 == null || cell1.getValue() == null) {
			return "";
		}
		String strtext = "";
		switch (cell1.getType()) {
		case NUMERIC:
			DecimalFormat df = getJavaDecimalFormatString(cell1.getCellstyle().getDataformat());
	        PubInfo.printStr(cell1.getCellstyle().getDataformat());
			strtext = df.format(cell1.getValue());

			break;
		case DATE:
			SimpleDateFormat ddf = getJavaDateFormatString(cell1.getCellstyle().getDataformat());
			strtext = ddf.format(cell1.getValue());
			break;
		default:
			strtext = cell1.getText();
			break;
		}
		return strtext;
	}

	private static DecimalFormat getJavaDecimalFormatString(String fmt) {
		String jfmt = "";
		int int1 = BuiltinFormats.getBuiltinFormat(fmt);
		switch (int1) {
		case 1:
		case 2:
		case 3:
		case 4:
		case 9:
		case 10:
			jfmt = fmt;
			break;
		case 5:
		case 6:
		case 42: // ￥90,987
			jfmt = "￥#,##0";
			break;
		case 7:
		case 8:
		case 44: // ￥90,987.80
			jfmt = "￥#,##0.00";
			break;
		case 11: // 4.56E13
			jfmt = "0.00E0";
			break;
		case 12:
		case 13:// 0.20
			jfmt = "#0.00";
			break;
		case 23:
		case 24:// $89,784
			jfmt = "$#,##0";
			break;
		case 25:
		case 26: // $3,456.90
			jfmt = "$#,##0.00";
			break;
		case 37:
		case 38: // 4,567
		case 41:
			jfmt = "#,##0";
			break;
		case 39:
		case 40:
		case 43: // 4,564.20
			jfmt = "#,##0.00";
			break;
		case 48:
			jfmt = "##0.0E0";
			break;
		case -1:
			jfmt = getJavaDecimalCustomFormatString(fmt);
			break;
		default:
			jfmt = "#0.######";
			break;
		}
		return new DecimalFormat(jfmt);
	}

	private static String getJavaDecimalCustomFormatString(String fmt) {
		String strOk = "#0.,$￥E";
		String jfmt = "";
		for (int i = 0; i < fmt.length(); i++) {
			String str1 = fmt.substring(i, i + 1);
			if (str1.equals(";")) {
				break;
			}
			if (strOk.indexOf(str1) >= 0) {
				jfmt += str1;
			}
		}
		return jfmt;
	}

	private static SimpleDateFormat getJavaDateFormatString(String fmt) {
		String jfmt = "";
		Locale lc = Locale.US;
		int int1 = BuiltinFormats.getBuiltinFormat(fmt);
		switch (int1) {
		case 14:
		case 30:
			jfmt = "d/M/yyyy";
			break;
		case 15:
			jfmt = "d-MMM-yyyy";
			break;
		case 16:
			jfmt = "d-MMM";
			break;
		case 17:
			jfmt = "MMM-yyyy";
			break;
		case 18:
			jfmt = "h:mm a";
			break;
		case 19:
			jfmt = "h:mm:ss a";
			break;
		case 20:
			jfmt = "h:mm";
			break;
		case 21:
			jfmt = "h:mm:ss";
			break;
		case 22:
			jfmt = "M/d/yyyy h:mm";
			break;
		case 27:
			jfmt = "yyyy年M月";
			break;
		case 28:
		case 29:
			jfmt = "M月d日";
			break;
		case 31:
			jfmt = "yyyy年M月d日";
			break;
		case 32:
			jfmt = "h时mm分";
			break;
		case 33:
			jfmt = "h时mm分ss秒";
			break;
		case 34:
			jfmt = "ah时mm分";
			lc = Locale.CHINESE;
			break;
		case 35:
			jfmt = "ah时mm分ss秒";
			lc = Locale.CHINESE;
			break;
		case 36:
			jfmt = "yyyy年M月";
			break;
		case 45:
		case 46:
		case 47:
			jfmt = "mm:ss";
			break;
		case -1:
			jfmt = getJavaDateCustomFormatString(fmt);
			if (fmt.indexOf("上午/下午") >= 0) {
				lc = Locale.CHINESE;
			}
			break;
		default:
			jfmt = "yyyy/M/d h:mm";
		}
		return new SimpleDateFormat(jfmt, lc);
	}

	private static String getJavaDateCustomFormatString(String fmt) {
		String jfmt = "";
		jfmt = fmt;
		jfmt = jfmt.replace("\"", "");
		jfmt = jfmt.replace("上午/下午", "a");
		jfmt = jfmt.replace("AM/PM", "a");
		jfmt = jfmt.replace("m", "M");
		jfmt = jfmt.replace("mmm", "MMM");
		return jfmt;
	}

	public static int getDecimalFormatDotcount(String fmt) {
		DecimalFormat df = getJavaDecimalFormatString(fmt);
        return df.getMinimumFractionDigits();
	}
}
