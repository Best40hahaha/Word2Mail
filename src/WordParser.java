import org.apache.poi.xwpf.usermodel.*;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * @author ：FlyingRedPig
 * @description：parse word template to hashmap
 * @date ：7/5/2020 10:53 AM
 */
public class WordParser {

	/*get the main content of the template, assume the first table is the template content**/
	public static XWPFTable getContent(String path) throws IOException {
		XWPFDocument doc = new XWPFDocument(new FileInputStream(new File(path)));
		List<XWPFTable> tables = doc.getTables();
		return tables.get(0);
	}

	/*helper method, to deal with the new line issue, we need to involve paragraph in my process, this method gives us the most basic
	* function to transfer multiple lines text to html format (<br> to build a new line) **/
	private static String paragraphsParser(List<XWPFParagraph> paragraphs){
		StringBuffer buffer = new StringBuffer();
		Iterator<XWPFParagraph> paragraphIterator = paragraphs.iterator();
		buffer.append(paragraphIterator.next().getText());
		while(paragraphIterator.hasNext()){
			buffer.append("<br>");
			buffer.append(paragraphIterator.next().getText());
		}
		return buffer.toString();
	}

	private static boolean isContainTable(XWPFTableCell cell){
		return cell.getTables().size() != 0;
	}

	/*helper method
	parse a row of a table to the html format
	* eg: input : | xuzinan | william | Tom |
	*     out put: <td>xuzinan</td><td>william</william><td>Tom</td>
	* **/
	private static String rowParser(XWPFTableRow row){
		StringBuffer buffer = new StringBuffer();
		List<XWPFTableCell> cells = row.getTableCells();
		for (XWPFTableCell cell : cells){
			buffer.append("<td>");
			// similarly, to avoid new line issue, call paragraphsParser
			List<XWPFParagraph> paragraphs = cell.getParagraphs();
			String text = paragraphsParser(paragraphs);
			buffer.append(text);
			buffer.append("</td>");

		}
		return buffer.toString();
	}

	/*helper method
	* parse a table in word (XWPFTable) to html format
	* **/
	private static String tableParser(XWPFTable table){
		StringBuffer buffer = new StringBuffer();
		List<XWPFTableRow> rows = table.getRows();
		Iterator<XWPFTableRow> rowIterator = rows.iterator();
		for (XWPFTableRow row : rows){
			buffer.append("<tr>");
			buffer.append(rowParser(row));
			buffer.append("</tr>");

		}
		buffer.insert(0,"<table>");
		buffer.append("</table>");
		return buffer.toString();
	}

	/*helper method, transfer the table to Hashmap, strong restrict here, it only read the first two columns of the table**/
	@NotNull
	private static HashMap<String, String> tableToMap(@NotNull XWPFTable table){
		HashMap<String, String> template = new HashMap<>();
		List<XWPFTableRow> rows = table.getRows();
		for(XWPFTableRow row : rows){
			List<XWPFTableCell> cells = row.getTableCells();
			XWPFTableCell cell1 = cells.get(0);
			XWPFTableCell cell2 = cells.get(1);
			String key = cell1.getText();
			List<XWPFParagraph> paragraphs = cell2.getParagraphs();
			String value = paragraphsParser(paragraphs);
			// should consider about that some cells contain tables
			if(isContainTable(cell2)){
				StringBuffer buffer = new StringBuffer();
				int tableNum = 0;
				Iterator<XWPFTable> tables = cell2.getTables().iterator();
				for(XWPFParagraph paragraph : paragraphs){
					String text = paragraph.getText();
					if(text.equals("")){
						String tableContent = tableParser(tables.next());
						buffer.append(tableContent);
					}else{
						buffer.append(text);
					}
				}
				value = buffer.toString();
			}
			template.put(key, value);
		}
		return template;
	}

	@NotNull
	public static HashMap<String, String> parse(String path) throws IOException {
		HashMap<String, String> reader = tableToMap(getContent(path));
		return reader;
	}

	public static void main(String[] args) throws IOException {
		String path = "C:\\Users\\C5293427\\Desktop\\test.docx";
		HashMap<String, String> h = WordParser.parse(path);
		System.out.println(h);
	}


}
