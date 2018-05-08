package acmr.excel;

import acmr.math.CalculateExpression;
import acmr.math.CalculateFunction;
import acmr.math.entity.MathException;
import acmr.util.PubInfo;

public class ExcelFunction extends CalculateFunction {
  
	
	public String function_aaddb(String a, String b) {
		return function_add(a, b);
	}
	public String checkfunction_aaddb(String a, String b) {
		
		return function_sub(a,b);
	}
	
 
	public static void main(String[] args) throws MathException {
		CalculateExpression ss = new acmr.math.CalculateExpression();
		ss.setFunctionclass(new ExcelFunction());
		String str1 = ss.Eval("aaddb(78,33)-76");
		PubInfo.printStr(str1);
	}
}
