package indi.hb.excelutils;

import java.io.IOException;
import java.util.TreeSet;
import java.util.regex.Pattern;

import org.apache.log4j.Level;

import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.NonNull;

/**
 * Excel解析成html的table元素
 * @author wanghb
 * @datetime 2018/11/06
 */
public class Excel2Table {
	private org.apache.log4j.Logger log = org.apache.log4j.Logger.getLogger(this.getClass().getName());
	/**
	 * 默认行高
	 */
	private final int DEFAULT_ROWS = 1;
	/**
	 * 默认列宽
	 */
	private final int DEFAULT_COLS = 1;
	
	/**
	 * Excel转Table代码
	 * 适用于本地环境
	 * @param excelPath 文件路径
	 * @return
	 */
	public String toTable(@NonNull String excelPath) {
		return toTable(readExcel(excelPath));
	}
	/**
	 * Excel转Table代码
	 * @param excel 文件
	 * @return
	 */
	public String toTable(@NonNull java.io.File excel) {
		return toTable(analyze(excel));
	}
	/**
	 * Excel转Table代码
	 * @param sheetBean
	 * @return
	 */
	public String toTable(@NonNull SheetBean sheetBean) {
		StringBuilder table = new StringBuilder();
		table.append("<table>\n");
		TreeSet<CellBean> cells = sheetBean.cells;
		// 每行中单元格的序号
		int cellNum = 0;
		// 上一个Cell的行号
		int lastX = 0;
		// 遍历单元格
		for (CellBean cellBean : cells) {
			// 矫正行判断
			if (cellBean.getY() > lastX) {
				lastX = cellBean.getY();
				cellNum = 0;
			}
			// 每行开始使用<tr>
			if (cellNum == 0) {
				table.append("<tr>\n");
				cellNum = cellBean.getX();
			}
			// 序号增加（单元格所占列数）
			cellNum += cellBean.getCols();
			table.append("<td").append((cellBean.getCols() > 1) ? " colspan=\"" + cellBean.getCols() + "\"" : "")
				.append((cellBean.getRows() > 1) ? " rowspan=\"" + cellBean.getRows() + "\"" : "").append(">");
			table.append(cellBean.getContent());
			table.append("</td>\n");
			// 每行结束，归零
			if (cellNum == sheetBean.getCols()) {
				table.append("</tr>\n");
				cellNum = 0;
			}
		}
		table.append("</table>\n");
		return table.toString();
	}
	/**
	 * 解析表格
	 * 默认只读取工作簿中的第一个工作表
	 * @param excel
	 * @return
	 */
	public SheetBean analyze(@NonNull java.io.File excel) {
		SheetBean sheetBean = new SheetBean();
		TreeSet<CellBean> cells = new TreeSet<CellBean>();
		Workbook workbook = null;
		try {
			workbook = Workbook.getWorkbook(excel);
		} catch (BiffException e) {
			log.log(Level.ERROR, new Throwable("文件类型错误！"));
			return null;
		} catch (IOException e) {
			log.log(Level.ERROR, new Throwable("文件读取失败！"));
			return null;
		}
		// 默认读取第一张工作表
		Sheet sheet = workbook.getSheet(0);
		// 合并的单元格
		Range[] mergeCells = sheet.getMergedCells();
		sheetBean.setRows(sheet.getRows());
		sheetBean.setCols(sheet.getColumns());
		CellBean[][] arrays = new CellBean[sheet.getColumns()][sheet.getRows()];
		CellBean cellBean = null;
		String content;
		for (int j, i = 0; i < sheet.getRows(); i++) {
			for (j = 0; j < sheet.getColumns(); j++) {
				content = sheet.getCell(j, i).getContents();
				// 空单元格暂不处理
				if (content != null && !content.isEmpty()) {
					cellBean = new CellBean(content, j, i, DEFAULT_ROWS, DEFAULT_COLS);
					arrays[j][i] = cellBean;
				}
				cells.add(cellBean);
			}
		}
		CellBean mergeCellBean = null;
		// 找出合并单元格,改写行高和列宽
		for (Range mergeCell : mergeCells) {
			mergeCellBean = arrays[mergeCell.getTopLeft().getColumn()][mergeCell.getTopLeft().getRow()];
			mergeCellBean.setCols(mergeCell.getBottomRight().getColumn() - mergeCell.getTopLeft().getColumn() + DEFAULT_COLS);
			mergeCellBean.setRows(mergeCell.getBottomRight().getRow() - mergeCell.getTopLeft().getRow() + DEFAULT_ROWS);
		}
		sheetBean.setCells(cells);
		return sheetBean;
	}
	/**
	 * 读取文件
	 * @param filePath
	 * @return
	 */
	public java.io.File readExcel(@NonNull String filePath) {
		java.io.File excel = null;
		String msg;
		if (!Pattern.matches(".*\\.xls$", filePath)) {
			msg = "Excel文件只支持xls类型！";
			log.log(Level.ERROR, msg, new Throwable(msg));
			return null;
		}
		excel = new java.io.File(filePath);
		return excel;
	}
	/**
	 * 单元格对象
	 * @author wanghb
	 */
	@Data
	@AllArgsConstructor
	@NoArgsConstructor
	public class CellBean implements java.util.Comparator<CellBean>, java.lang.Comparable<CellBean> {
		/**
		 * 内容
		 */
		private String content;
		/**
		 * 列号
		 */
		private int x;
		/**
		 * 行号
		 */
		private int y;
		/**
		 * 占据行数
		 */
		private int rows;
		/**
		 * 占据列数
		 */
		private int cols;
		/**
		 * 按照单元格位置排序，左上<右下
		 */
		public int compare(CellBean o1, CellBean o2) {
			int i = 0;
			if (o1.y > o2.y) {
				// 1.行号越大越靠后
				i = -1;
			} else if (o1.y < o2.y) {
				// 2.行号越小越靠前
				i = 1;
			} else if (o1.x > o2.x) {
				// 3.行号相同，比列号，列号越大越靠后
				i = -1;
			} else if (o1.x < o2.x) {
				// 4.列号越小越靠前
				i = 1;
			}
			// 5.行号、列号均相同则为相同单元格
			return i;
		}
		/**
		 * 降序排列
		 */
		public int compareTo(CellBean o) {
			return compare(o, this);
		}
	}
	/**
	 * 工作表对象
	 * @author wanghb
	 */
	@Data
	public class SheetBean {
		/**
		 * 总行数
		 */
		private int rows;
		/**
		 * 总列数
		 */
		private int cols;
		/**
		 * 单元格序列
		 */
		private TreeSet<CellBean> cells;
	}
}
