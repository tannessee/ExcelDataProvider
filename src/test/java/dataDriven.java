import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public ArrayList<String> getData(String testcaseName) throws IOException {

		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("C://Udemy//demodata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets(); // количсетво страниц

		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {

				XSSFSheet sheet = workbook.getSheetAt(i);

				// Identify TestCases column by scanning 1st row
				Iterator<Row> rows = sheet.iterator(); // sheet is collection of rows
				Row firstrow = rows.next();

				Iterator<Cell> ce = firstrow.cellIterator(); // row is collection of cells
				// int k = 0;
				int column = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					//System.out.print(value.getStringCellValue() + ' ');
					if (value.getStringCellValue().equalsIgnoreCase("testcases")) {

						column = value.getColumnIndex(); // индекс колонки
					}
					// k++;
				}
				System.out.println();
				// System.out.println(column);
				// When column is identified then scan entire testcase column to find PUrchase
				while (rows.hasNext()) {
					Row r = rows.next(); // rows -итератор строк как i

					if (r.getCell(column) != null && r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						// System.out.println(r.getCell(column));

						// After grabbing purchase testcase row = pull all the data of that row and feed
						// into test
						Iterator<Cell> cv = r.cellIterator(); // cv - итератор ячеек

						while (cv.hasNext()) {
							Cell c = cv.next();
							if (c.getCellType() == CellType.NUMERIC) {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							} else {
								a.add(c.getStringCellValue());

							}

						}
					}

				}

			}

		}
		return a;

	}

	public static void main(String[] args) throws IOException {

	}

}
