/*** Eclipse Class Decompiler plugin, copyright (c) 2012 Chao Chen (cnfree2000@hotmail.com) ***/
package acmr.excel.pojo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;

public class ExcelDateUtil {
	public static boolean isCellDateFormatted(Cell cell) {
		if (cell == null)
			return false;
		boolean bDate = false;

		double d = cell.getNumericCellValue();
		if (DateUtil.isValidExcelDate(d)) {
			CellStyle style = cell.getCellStyle();
			if (style == null)
				return false;
			int i = style.getDataFormat();
			String f = style.getDataFormatString();
			bDate = isADateFormat(i, f);
		}
		return bDate;
	}

	private static boolean isADateFormat(int i, String f) {
		if (f != null) {
			f = f.replaceAll("[\"|\']", "").replaceAll("[年|月|日|时|分|秒|毫秒|微秒]", "");
		}
		boolean mark1 = DateUtil.isADateFormat(i, f);
		return mark1;
	}

}