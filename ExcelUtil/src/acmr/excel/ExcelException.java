package acmr.excel;

/**
 * 操作excel包返回的错误
 * @author zengqu
 *
 */
public class ExcelException extends Exception {

	private static final long serialVersionUID = 1L;

	public ExcelException() {
		super();
	}

	public ExcelException(String e) {
		super(e);
	}

	public ExcelException(Exception e) {
		super(e);
	}

	public ExcelException(String e, Exception e2) {
		super(e, e2);
	}
}
