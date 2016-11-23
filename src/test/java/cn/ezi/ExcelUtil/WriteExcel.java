package cn.ezi.ExcelUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class WriteExcel {
	@Test
	public void create() {
		FileOutputStream out =null;
		XSSFWorkbook wb = null;
		String[][] data = { { "angelabady", "跑了", "跑了", "跑了", "跑了" }, { "邓超", "跑了", "跑了", "跑了", "跑了" },
				{ "网租赁", "跑了", "跑了", "跑了", "跑了" }, { "郑楷", "跑了", "跑了", "跑了", "跑了" } };
		String[] titles = { "姓名", "Monday", "Tuesday", "Wednesday", "Ruursday", "Friday" };
		try {
			File file = new File("E:/new.xlsx");
			wb = new XSSFWorkbook();
			Sheet sheet = wb.createSheet();
			Row handerRow = sheet.createRow(0);
			handerRow.setHeightInPoints(12.75f);
			for (int i = 0; i < titles.length; i++) {
				Cell cell = handerRow.createCell(i);
				cell.setCellValue(titles[i]);
			}
			Row row;
			Cell cell;
			int rownum = 1;
			for (int i = 0; i < data.length; i++, rownum++) {
				row = sheet.createRow(rownum);
				if (data[i] == null)
					continue;
				for (int j = 0; j < data[i].length; j++) {
					
					cell = row.createCell(j);
					cell.setCellValue(data[i][j]);
					
				}
			}	
			out = new FileOutputStream(file);
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			if(out!=null){
				try {
					out.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			if(wb!=null){
				try {
					wb.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

	}
}
