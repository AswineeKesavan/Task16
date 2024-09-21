package Task16;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Question3 {

	public static void main(String[] args) {
		Question3 x = new Question3();

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");

		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] { "Name", "Age", "Email" });
		data.put("2", new Object[] { "John Doe", 30, "john@test.com" });
		data.put("3", new Object[] { "Jane Doe", 28, "john@test.com" });
		data.put("4", new Object[] { "Bob Smith", 35, "jacky@example.com" });
		data.put("5", new Object[] { "Swapnil", 37, "swapnil@example.com" });

		// Iterate over data and write to sheet

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {

			XSSFRow row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}

		// Write the workbook in file system
		try {
			FileOutputStream out = new FileOutputStream(
					new String("C:\\Users\\ramak\\eclipse-workspace\\GuviTask\\src\\Task16\\Task16.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}

		// read the workbook in file system
		for (int i = 0; i < 5; i++) {
			for (int j = 0; j < 3; j++) {
				String data1 = x.getExcelData("Sheet1", i, j);
				System.out.println(data1);
			}
		}
	}

	public String getExcelData(String sheetName, int rowNum, int colNum) {
		String retVal = null;

		try {
			FileInputStream fis = new FileInputStream(
					"C:\\Users\\ramak\\eclipse-workspace\\GuviTask\\src\\Task16\\Task16.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet s = wb.getSheet(sheetName);
			XSSFRow r = s.getRow(rowNum);
			XSSFCell c = r.getCell(colNum);
			retVal = Question3.getCellValue(c);
			fis.close();
			wb.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return retVal;

	}

	public static String getCellValue(XSSFCell c) {
		switch (c.getCellType()) {
		case NUMERIC:
			return String.valueOf(c.getNumericCellValue()); // 10 -> "10"
		case BOOLEAN:
			return String.valueOf(c.getBooleanCellValue());
		case STRING:
			return c.getStringCellValue();
		default:
			return c.getStringCellValue();
		}
	}
}
