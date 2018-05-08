/**
 * Xls2Xlsx.java
 *北京华通人商用信息有限公司版权所有
 */
package acmr.excel;

import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.PaneInformation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.util.PubInfo;

/**
 * 负责讲xsl格式的Excel转换成XLSX格式
 * 
 * @author zhangqiang
 * 
 */
public class Xls2Xlsx {

	private int lastColumn = 0;
	private HashMap<Integer, XSSFCellStyle> styleMap = new HashMap<Integer, XSSFCellStyle>();

	public void transformHSSF(HSSFWorkbook workbookOld, XSSFWorkbook workbookNew) {
		XSSFSheet sheetNew;
		HSSFSheet sheetOld;
		workbookNew.setMissingCellPolicy(workbookOld.getMissingCellPolicy());

		for (int i = 0; i < workbookOld.getNumberOfSheets(); i++) {
			sheetOld = workbookOld.getSheetAt(i);
			sheetNew = workbookNew.getSheet(sheetOld.getSheetName());
			sheetNew = workbookNew.createSheet(sheetOld.getSheetName());
			this.transform(workbookOld, workbookNew, sheetOld, sheetNew);
		}
	}

	private void transform(HSSFWorkbook workbookOld, XSSFWorkbook workbookNew, HSSFSheet sheetOld, XSSFSheet sheetNew) {

		sheetNew.setDisplayFormulas(sheetOld.isDisplayFormulas());
		sheetNew.setDisplayGridlines(sheetOld.isDisplayGridlines());
		sheetNew.setDisplayGuts(sheetOld.getDisplayGuts());
		sheetNew.setDisplayRowColHeadings(sheetOld.isDisplayRowColHeadings());
		sheetNew.setDisplayZeros(sheetOld.isDisplayZeros());
		sheetNew.setFitToPage(sheetOld.getFitToPage());
		//
		// TODO::sheetNew.setForceFormulaRecalculation(sheetOld.getForceFormulaRecalculation());
		sheetNew.setHorizontallyCenter(sheetOld.getHorizontallyCenter());
		sheetNew.setMargin(Sheet.BottomMargin, sheetOld.getMargin(Sheet.BottomMargin));
		sheetNew.setMargin(Sheet.FooterMargin, sheetOld.getMargin(Sheet.FooterMargin));
		sheetNew.setMargin(Sheet.HeaderMargin, sheetOld.getMargin(Sheet.HeaderMargin));
		sheetNew.setMargin(Sheet.LeftMargin, sheetOld.getMargin(Sheet.LeftMargin));
		sheetNew.setMargin(Sheet.RightMargin, sheetOld.getMargin(Sheet.RightMargin));
		sheetNew.setMargin(Sheet.TopMargin, sheetOld.getMargin(Sheet.TopMargin));
		sheetNew.setPrintGridlines(sheetNew.isPrintGridlines());
		// TODO::sheetNew.setRightToLeft(sheetNew.isRightToLeft());
		sheetNew.setRowSumsBelow(sheetNew.getRowSumsBelow());
		sheetNew.setRowSumsRight(sheetNew.getRowSumsRight());
		sheetNew.setVerticallyCenter(sheetOld.getVerticallyCenter());

		XSSFRow rowNew;
		for (Row row : sheetOld) {
			rowNew = sheetNew.createRow(row.getRowNum());
			if (rowNew != null)
				this.transform(workbookOld, workbookNew, sheetOld, (HSSFRow) row, rowNew);
		}

		for (int i = 0; i < this.lastColumn; i++) {
			sheetNew.setColumnWidth(i, sheetOld.getColumnWidth(i));
			sheetNew.setColumnHidden(i, sheetOld.isColumnHidden(i));
		}

		for (int i = 0; i < sheetOld.getNumMergedRegions(); i++) {
			CellRangeAddress merged = sheetOld.getMergedRegion(i);
			sheetNew.addMergedRegion(merged);
		}
		PaneInformation paninfo = sheetOld.getPaneInformation();// 有冻结设置
		if (paninfo != null) {
			int row = paninfo.getHorizontalSplitTopRow();
			int col = paninfo.getVerticalSplitLeftColumn();
			if (row > 0 && col > 0) {
				sheetNew.createFreezePane(col, row);
			}
		}
	}

	private void transform(HSSFWorkbook workbookOld, XSSFWorkbook workbookNew, HSSFSheet sheetOld, HSSFRow rowOld, XSSFRow rowNew) {
		XSSFCell cellNew;
		short rowheight = rowOld.getHeight();
		short defautltrowheight = sheetOld.getDefaultRowHeight();
		if (defautltrowheight != rowheight) {
			rowNew.setHeight(rowheight);
		}
		// TODO::if (rowOld.getRowStyle() != null) {
		/*
		 * Integer hash = rowOld.getRowStyle().hashCode(); if
		 * (!this.styleMap.containsKey(hash))
		 * this.transform(workbookOld,workbookNew,hash,
		 * (XSSFCellStyle)rowOld.getRowStyle
		 * (),(HSSFCellStyle)workbookNew.createCellStyle());
		 * rowNew.setRowStyle(this.styleMap.get(hash)); }
		 */
		for (Cell cell : rowOld) {
			cellNew = rowNew.createCell(cell.getColumnIndex(), cell.getCellType());
			if (cellNew != null)
				this.transform(workbookOld, workbookNew, (HSSFCell) cell, cellNew);
		}
		this.lastColumn = Math.max(this.lastColumn, rowOld.getLastCellNum());
	}

	private void transform(HSSFWorkbook workbookOld, XSSFWorkbook workbookNew, HSSFCell cellOld, XSSFCell cellNew) {
		cellNew.setCellComment(cellOld.getCellComment());

		Integer hash = cellOld.getCellStyle().hashCode();
		if (this.styleMap != null && !this.styleMap.containsKey(hash)) {
			this.transform(workbookOld, workbookNew, hash, cellOld.getCellStyle(), (XSSFCellStyle) workbookNew.createCellStyle());
		}
		cellNew.setCellStyle(this.styleMap.get(hash));
		switch (cellOld.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			cellNew.setCellValue(cellOld.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			cellNew.setCellValue(cellOld.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			try {
				cellNew.setCellFormula(cellOld.getCellFormula());
				if (cellOld.getCachedFormulaResultType() == 0) {
					cellNew.setCellValue(cellOld.getNumericCellValue());
				} else {
					cellNew.setCellValue(cellOld.getStringCellValue());
				}
			} catch (Exception e) {
				cellNew.setCellType(Cell.CELL_TYPE_BLANK);
				e.printStackTrace();
			}
			break;
		case Cell.CELL_TYPE_NUMERIC:
			cellNew.setCellValue(cellOld.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING:
			String cellvalue = cellOld.getStringCellValue();
			if (cellvalue != null && !cellvalue.isEmpty()) {
				RichTextString newrichtextstr = transform(workbookOld, workbookNew, cellOld.getRichStringCellValue());
				if (newrichtextstr != null) {
					cellNew.setCellValue(newrichtextstr);
				} else {
					cellNew.setCellValue(cellvalue);
				}
			}
			break;
		default:
		}
	}

	private RichTextString transform(Workbook workbookOld, Workbook newworkbook, RichTextString oldrichtextstr) {
		RichTextString newrichtextstr = new XSSFRichTextString(oldrichtextstr.getString());
		int len = oldrichtextstr.length();
		boolean hasfont = false;
		HSSFRichTextString hssfoldrichtextstr = (HSSFRichTextString) oldrichtextstr;
		for (int i = 0; i < len; i++) {
			short fontindex = hssfoldrichtextstr.getFontAtIndex(i);
			if (fontindex > 0) {
				XSSFFont newfont = transform((XSSFWorkbook) newworkbook, (HSSFFont) workbookOld.getFontAt(fontindex));
				if (newfont != null) {
					hasfont = true;
					newrichtextstr.applyFont(i, i + 1, newfont);
				}
			}
		}

		if (!hasfont) {
			return null;
		}
		return newrichtextstr;
	}

	private void transform(HSSFWorkbook workbookOld, XSSFWorkbook workbookNew, Integer hash, HSSFCellStyle styleOld, XSSFCellStyle styleNew) {
		styleNew.setAlignment(styleOld.getAlignment());
		try {
			styleNew.setBorderBottom(styleOld.getBorderBottom());
		} catch (Exception e) {
			PubInfo.printStr("设置下边框时出错");
		}
		try {
			styleNew.setBorderLeft(styleOld.getBorderLeft());
		} catch (Exception e) {
			PubInfo.printStr("设置左边框时出错");
		}
		try {
			styleNew.setBorderRight(styleOld.getBorderRight());
		} catch (Exception e) {
			PubInfo.printStr("设置右边框时出错");
		}
		try {
			styleNew.setBorderTop(styleOld.getBorderTop());
		} catch (Exception e) {
			PubInfo.printStr("设置上边框时出错");
		}
		try {
			styleNew.setDataFormat(this.transform(workbookOld, workbookNew, styleOld.getDataFormatString(), styleOld.getDataFormat()));
		} catch (Exception e) {
			PubInfo.printStr("setDataFormat时出错");
		}
		if (styleOld.getFillBackgroundColor() != HSSFColor.AUTOMATIC.index) {
			try {
				styleNew.setFillBackgroundColor(styleOld.getFillBackgroundColor());
			} catch (Exception e) {
				PubInfo.printStr("复制格式时出错");
			}
		}
		if (styleOld.getFillForegroundColor() != HSSFColor.AUTOMATIC.index) {
			try {
				styleNew.setFillForegroundColor(styleOld.getFillForegroundColor());
				styleNew.setFillPattern(styleOld.getFillPattern());
			} catch (Exception e) {
				PubInfo.printStr("复制格式时出错");
			}
		}
		XSSFFont newfont = transform(workbookNew, styleOld.getFont(workbookOld));
		if (newfont != null) {
			try {
				styleNew.setFont(newfont);
			} catch (Exception e) {
				PubInfo.printStr("复制格式时出错");
			}
		}
		try {
			styleNew.setHidden(styleOld.getHidden());
			styleNew.setIndention(styleOld.getIndention());
			styleNew.setLocked(styleOld.getLocked());
			styleNew.setVerticalAlignment(styleOld.getVerticalAlignment());
			styleNew.setWrapText(styleOld.getWrapText());
		} catch (Exception e) {
			PubInfo.printStr("复制格式时出错");
		}
		this.styleMap.put(hash, styleNew);
	}

	private short transform(Workbook workbookOld, Workbook newworkbook, String dataFormat, short format) {
		DataFormat formatNew = newworkbook.createDataFormat();
		if (dataFormat == null || dataFormat.isEmpty()) {
			return format;
		}
		return formatNew.getFormat(dataFormat);
	}

	private XSSFFont transform(XSSFWorkbook workbookNew, HSSFFont fontOld) {
		if (fontOld == null || fontOld.getIndex() < 0) {
			return null;
		}
		XSSFFont fontNew = workbookNew.createFont();
		fontNew.setBoldweight(fontOld.getBoldweight());
		fontNew.setCharSet(fontOld.getCharSet());
		fontNew.setColor(fontOld.getColor());
		fontNew.setFontName(fontOld.getFontName());
		fontNew.setFontHeight(fontOld.getFontHeight());
		fontNew.setItalic(fontOld.getItalic());
		fontNew.setStrikeout(fontOld.getStrikeout());
		fontNew.setTypeOffset(fontOld.getTypeOffset());
		fontNew.setUnderline(fontOld.getUnderline());
		return fontNew;
	}

}