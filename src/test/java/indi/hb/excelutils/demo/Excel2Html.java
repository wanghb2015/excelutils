package indi.hb.excelutils.demo;

import java.awt.Desktop;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URI;

import org.apache.log4j.Level;

import indi.hb.excelutils.Excel2Table;
import lombok.Cleanup;
import lombok.NonNull;

public class Excel2Html {
	private org.apache.log4j.Logger log = org.apache.log4j.Logger.getLogger(this.getClass().getName());
	/**
	 * 缺省取第一个入参
	 * @param args
	 */
	public static void main(String[] args) {
		String filePath = "C:\\Users\\wangh\\Desktop\\test.xls";
		if (args.length > 0 && args[0] != null && !args[0].isEmpty()) {
			filePath = args[0];
		}
		Excel2Html excel2Html = new Excel2Html();
		excel2Html.preview(filePath);
	}
	/**
	 * 调用本地默认浏览器预览生成的HTML
	 * @param filePath
	 */
	public void preview(@NonNull String filePath) {
		String code = toHtml(filePath);
		if (code == null || code.isEmpty()) {
			log.log(Level.ERROR, new Throwable("解析Excel失败！"));
		}
		//控制台输出
		System.out.println(code);
		//生成html文件
		File html = new File("test.html");
		if (!html.exists()) {
			try {
				html.createNewFile();
			} catch (IOException e) {
				log.log(Level.ERROR, new Throwable("创建测试页失败，请自行复制控制台html代码！"));
			}
		}
		//写入html代码
		try {
			@Cleanup FileOutputStream fos = new FileOutputStream(html);
			fos.write(code.getBytes("UTF-8"));
		} catch (FileNotFoundException e) {
			log.log(Level.ERROR, new Throwable("测试页意外丢失，请自行复制控制台html代码！"));
		} catch (UnsupportedEncodingException e) {
			log.log(Level.DEBUG, new Throwable("不支持的编码格式"));
		} catch (IOException e) {
			log.log(Level.ERROR, new Throwable("测试页写入失败，请自行复制控制台html代码！"));
		}
		//调用系统默认浏览器预览生成的HTML
		if (Desktop.isDesktopSupported()) {
			URI uri = URI.create("test.html");
			Desktop desktop = Desktop.getDesktop();
			if (desktop.isSupported(Desktop.Action.BROWSE)) {
				try {
					desktop.browse(uri);
				} catch (IOException e) {
					log.log(Level.WARN, new Throwable("获取不到系统默认浏览器"));
				}
			}
		}
	}
	/**
	 * Excel转换成Table代码，并生成Html代码
	 * 最后一行增加下拉框
	 * @param filePath 本地文件路径
	 * @return
	 */
	public String toHtml(@NonNull String filePath) {
		StringBuilder html = new StringBuilder();
		Excel2Table excel2Table = new Excel2Table();
		String table = excel2Table.toTable(filePath);
		if (table != null && !table.isEmpty()) {
			html.append("<!doctype html>\n" + 
					"<html>\n" + 
					"<head>\n" + 
					"<meta charset=\"utf-8\" />\n" + 
					"<style type=\"text/css\">\n" + 
					"table {\n" + 
					"    border-collapse: collapse;\n" + 
					"    border: none;\n" + 
					"    margin: 0 auto;\n" + 
					"}\n" + 
					"td {\n" + 
					"    border: solid #000 1px;\n" + 
					"}\n" + 
					"</style>\n" + 
					"</head>\n" + 
					"<body>\n");
			html.append(table);
			html.append("</body>\n" + 
					"<script type=\"text/javascript\">\n" + 
					"window.onload = function() {\n" + 
					"	addRow();\n" + 
					"}\n" + 
					"function addRow() {\n" + 
					"	//缺省取第一个table元素\n" + 
					"	var tab = document.getElementsByTagName(\"table\")[0];\n" + 
					"	//行,单元格\n" + 
					"	var row, cell;\n" + 
					"	//列数,总行数,单元格行高\n" + 
					"	var cnt_col, cnt_row = tab.rows.length;\n" + 
					"    //标签名，二维数组\n" + 
					"    var newTagNames = new Array(new Array());\n" + 
					"    //遍历第一行，确定二维数组x轴，之后填充空格\n" + 
					"    for (var j = 0; j < tab.rows.length; j++) {\n" + 
					"    	row = tab.rows[j];\n" + 
					"    	cnt_col = 0;\n" + 
					"		for (var i = 0; i < row.cells.length; i++) {\n" + 
					"			cell = row.cells[i];\n" + 
					"			for (var m = 0; m < cell.rowSpan; m++) {\n" + 
					"				for (var n = 0; n < cell.colSpan; n++) {\n" + 
					"					fillArray(newTagNames, n + cnt_col, m + j, cell.innerHTML);\n" + 
					"				}\n" + 
					"			}\n" + 
					"	        cnt_col += cell.colSpan;\n" + 
					"		}\n" + 
					"    }\n" + 
					"	//插入新列\n" + 
					"	var newCell, newRow = tab.insertRow(cnt_row);\n" + 
					"	for (var k = 0; k < newTagNames[newTagNames.length - 1].length; k++) {\n" + 
					"		newCell = newRow.insertCell(k);\n" + 
					"		newCell.innerHTML = createSelect(k + 1, newTagNames[newTagNames.length - 1][k]);\n" + 
					"	}\n" + 
					"}\n" + 
					"/**\n" + 
					" * 数据填充\n" + 
					" */\n" + 
					"function fillArray(a, x, y, c) {\n" + 
					"	if (typeof a != \"object\" && !a instanceof Array) {\n" + 
					"		//类型不正确\n" + 
					"		return;\n" + 
					"	}\n" + 
					"	if (y >= a.length) {\n" + 
					"		a.push(new Array());\n" + 
					"		fillArray(a, x, y, c);\n" + 
					"	} else if (x >= a[y].length) {\n" + 
					"		a[y].push(\"\");\n" + 
					"        fillArray(a, x, y, c);\n" + 
					"	} else if (a[y][x] == \"\") {\n" + 
					"		a[y][x] = c;\n" + 
					"	} else if (x < a[y].length) {\n" + 
					"		fillArray(a, x + 1, y, c);\n" + 
					"	}\n" + 
					"}\n" + 
					"function createSelect(i, n) {\n" + 
					"	var selectTag = \"<select id=\\\"value\" + i + \"\\\" name=\\\"\" + n + \"\\\">\"\n" + 
					"    + \"<option value=\\\"1\\\">文本</option>\"\n" + 
					"    + \"<option value=\\\"2\\\">日期</option>\"\n" + 
					"    + \"<option value=\\\"3\\\">数字</option>\"\n" + 
					"    + \"<option value=\\\"4\\\">选项</option>\"\n" + 
					"    + \"</select>\";\n" + 
					"    return selectTag;\n" + 
					"} \n" + 
					"</script>\n" + 
					"</html>");
		}
		return html.toString();
	}
}
