/**
 * 
 */
package excel.test;

import java.io.IOException;

import com.wxct.cxzx.excel.Excel2007;

/**
 * 2018-11-16 17:13:56
 * @author zhenr
 *
 */
public class Test {
	@org.junit.Test
	public void test() throws IOException {
		new Excel2007().readSheet("e://x.xlsx",0);
	}
}
