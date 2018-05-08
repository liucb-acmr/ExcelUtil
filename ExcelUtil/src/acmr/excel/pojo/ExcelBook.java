package acmr.excel.pojo;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.serializer.SerializerFeature;

import acmr.excel.ExcelException;
import acmr.excel.Xls2Xlsx;
import acmr.excel.Xlsx2Xls;
import acmr.excel.pojo.Constants.XLSTYPE;
import acmr.util.ListHashMap;
import acmr.util.PubInfo;

/**
 * @author zengqu 内存中保存的excelbook模型，
 */
public class ExcelBook implements Cloneable, Serializable {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private ListHashMap<ExcelSheet> sheets;

	private Map<String, String> exps;
	
	private List<Sheet> nativeSheet;

	/**
	 * 构造函数，初始化对象
	 */
	public ExcelBook() {
		sheets = new ListHashMap<ExcelSheet>();
		exps = new HashMap<String, String>();
		nativeSheet = new ArrayList<Sheet>();
	}

	@Override
	public ExcelBook clone() {
		ExcelBook o = new ExcelBook();
		for (int i = 0; i < sheets.size(); i++) {
			o.sheets.add(sheets.get(i).clone());
		}
		for (String key1 : exps.keySet()) {
			o.exps.put(key1, exps.get(key1));
		}
		return o;
	}

	/**
	 * 得到附加属性map表
	 * 
	 * @return map表
	 * @see Map
	 */
	public Map<String, String> getExps() {
		return exps;
	}

	public void setExps(Map<String, String> exps) {
		this.exps = exps;
	}

	/**
	 * 得到sheet集合
	 * 
	 * @return
	 * 
	 * @return list集合
	 * @see List
	 */
	public ListHashMap<ExcelSheet> getSheets() {
		return sheets;
	}
	

	public List<Sheet> getNativeSheet() {
		return nativeSheet;
	}

	public void setNativeSheet(List<Sheet> nativeSheet) {
		this.nativeSheet = nativeSheet;
	}

	/**
	 * 加载excel，把poi workbook 转成内存模型
	 * 
	 * @param b1
	 *            poi的excel 2007对象
	 * @see XSSFWorkbook
	 */
	public void LoadExcel(XSSFWorkbook b1) {
		sheets.clear();
		nativeSheet.clear();
		for (int i = 0; i < b1.getNumberOfSheets(); i++) {
			XSSFSheet s1 = b1.getSheetAt(i);
			PubInfo.printStr(s1.getSheetName());
			ExcelSheet sheet1 = new ExcelSheet();
			sheet1.loadSheet(s1);
			sheets.add(sheet1);
			nativeSheet.add(s1);
		}
	}

	/**
	 * 直接从文件中加载excel，会自动把2003格式转成2007的格式在加载
	 * 
	 * @param file
	 * @throws ExcelException
	 */
	public void LoadExcel(String file) throws Exception {
		try {
			FileInputStream fi = new FileInputStream(file);
			XLSTYPE type = XLSTYPE.XLSX;
			if (file.toLowerCase().endsWith(".xls")) {
				type = XLSTYPE.XLS;
			}
			LoadExcel(fi, type);
			fi.close();

		} catch (Exception e) {
			throw new ExcelException("加载没有成功!", e);
		}
	}

	/**
	 * 输入流中加载excel对象，需要指定excel格式
	 * 
	 * @param in
	 *            输入流，可是上传的文件流
	 * @param type
	 *            excel类型
	 * @throws ExcelException
	 */
	public void LoadExcel(InputStream in, XLSTYPE type) throws ExcelException {
		try {
			XSSFWorkbook b1 = null;
			if (type == XLSTYPE.XLS) {
				HSSFWorkbook book2 = new HSSFWorkbook(in);
				b1 = new XSSFWorkbook();
				new Xls2Xlsx().transformHSSF(book2, b1);
			} else {
				b1 = new XSSFWorkbook(in);
			}
			this.LoadExcel(b1);
			b1 = null;
		} catch (Exception e) {
			throw new ExcelException("加载没有成功!", e);
		}
	}

	/**
	 * 保存到输出流中，2003 或者2007格式 xls xlsx
	 * 
	 * @param out1
	 * @param type
	 * @throws ExcelException
	 */
	public void saveExcel(OutputStream out1, XLSTYPE type) throws ExcelException {
		try {
			XSSFWorkbook book1 = new XSSFWorkbook();
			List<ExcelCellStyle> cells = new ArrayList<ExcelCellStyle>();
			List<ExcelFont> fonts = new ArrayList<ExcelFont>();
			int intactive = -1;
			for (int i = 0; i < sheets.size(); i++) {
				XSSFSheet s1 = book1.createSheet();
				sheets.get(i).SaveToExcelSheet(s1, cells, fonts);
				if (intactive < 0 && sheets.get(i).getHiddenstate() == 0) {
					intactive = i;
				}
			}
			if (intactive >= 0) {// 确保当前sheet不是隐藏的sheet
				book1.setActiveSheet(intactive);
			}
			if (type == XLSTYPE.XLS) {
				HSSFWorkbook book2 = new HSSFWorkbook();
				new Xlsx2Xls().transformXSSF(book1, book2);
				if (intactive >= 0) {
					book2.setActiveSheet(intactive);
				}
				book2.write(out1);
			} else {
				book1.write(out1);
			}
		} catch (Exception e) {
			throw new ExcelException("保存没有成功!", e);
		}
	}

	/**
	 * 保存到文件中
	 * 
	 * @param file
	 * @throws ExcelException
	 */
	public void saveExcel(String file) throws ExcelException {
		try {
			FileOutputStream fo = new FileOutputStream(file);
			XLSTYPE type = XLSTYPE.XLSX;
			if (file.toLowerCase().endsWith(".xls")) {
				type = XLSTYPE.XLS;
			}
			saveExcel(fo, type);
			fo.close();
		} catch (Exception e) {
			throw new ExcelException("保存没有成功!", e);
		}
	}

	public byte[] SerializeBytes() {
		return PubInfo.getSerializeBytes(this);
	}

	public static ExcelBook SerializeObject(byte[] bs) {
		return (ExcelBook) PubInfo.getSerializeObject(bs);
	}

	public String JSONString() {
		return JSON.toJSONString(this, SerializerFeature.WriteClassName);
	}

	public static ExcelBook JSONParse(String str1) {
		ExcelBook book1 = JSON.parseObject(str1, ExcelBook.class);
		for (int i = 0; i < book1.getSheets().size(); i++) {
			book1.getSheets().get(i).afterJOSN();
		}
		return book1;
	}

	public static ExcelBook JSONParse(byte str1[]) {
		String str = "{}";
		if (null != str1 && str1.length > 0) {
			try {
				str = new String(str1, "utf-8");
			} catch (UnsupportedEncodingException e) {
				e.printStackTrace();
			}
		}
		return JSONParse(str);
	}

	public static void main(String[] args) throws Exception {

		String file = "d:/a.xlsx";
		ExcelBook book1 = new ExcelBook();
		PubInfo.printStr("1");
		book1.LoadExcel(file);
		ExcelSheet sheet = book1.getSheets().get(0);
		
		for(int i=0;i<sheet.getRows().size();i++){
			ExcelRow row = sheet.getRows().get(i);
			for(int j=0;j<row.getCells().size();j++){
				ExcelCell cell = row.getCells().get(j);
				System.out.println();
			}
		}

		/*PubInfo.printStr("2");
        book1.saveExcel("d:/a1.xlsx");
       // book1.getSheets().get(0).getRows().get(10).getCells().get(0).getShowText();
		String str1 = book1.JSONString();
		// PubInfo.printStr("" + str1);
		ExcelBook book2 = ExcelBook.JSONParse(str1);
		book2.saveExcel("d:/a11.xlsx");
		PubInfo.printStr("3");
		ExcelBook book3 = (ExcelBook) acmr.util.PubInfo.deepclone(book1);
		book3.saveExcel("d:/a13.xlsx");*/
	}

}
