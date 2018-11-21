/**
 * 
 */
package excel.test;

import java.io.IOException;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import com.wxct.cxzx.excel.Excel2007;

/**
 * 2018-11-16 17:13:56
 * @author zhenr
 *
 */
public class Test {
	@org.junit.Test
	public void test() throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
		
		List<List<Object>> list=new Excel2007().readSheet("e://a.xlsx",0);
		for(List<Object> rowValue : list) {
			for(Object cell : rowValue) {
				System.out.println(cell);
			}
		}
	}
}
