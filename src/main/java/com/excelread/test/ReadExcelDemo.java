package com.excelread.test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo {

	public static int getCellType(Cell cell) {
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
			return 0;
		else
			return 1;
	}

	public static String getMonth(String eDate) {
		String tokens[] = eDate.split("\\s");
		return tokens[1];
	}

	public static String getTime(String eDate) {
		String tokens[] = eDate.split("\\s");
		return tokens[3];
	}

	public static void main(String[] args) {
		try {
			int dayShifts = 0;
			int nightShifts = 0;

			BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(System.in));

			System.out.println("Enter the empNo");
			int empNo = Integer.valueOf(bufferedReader.readLine());
			System.out.println("Enter the Month");
			String empMonth = bufferedReader.readLine();
			FileInputStream file = new FileInputStream(new File("excel_final.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();

			List<Employee> employees = new ArrayList<Employee>();
			Employee employee;
			for (int i = 0; rowIterator.hasNext(); i++) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				employee = new Employee();
				boolean isEmpIdMatches = false;
				boolean isEmpMonthMatches = false;

				for (int j = 0; cellIterator.hasNext(); j++) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly

					if (i != 0) {
						switch (j) {
						case 0:
							if (ReadExcelDemo.getCellType(cell) == 0) {
								employee.setEmpid(Math.round(cell.getNumericCellValue()));
							} else {
								employee.setEmpid(Long.valueOf(cell.getStringCellValue()));
							}
							if (empNo == employee.getEmpid()) {
								System.out.println(empNo);
								isEmpIdMatches = true;
							}
							break;
						case 1:
							if (ReadExcelDemo.getCellType(cell) == 0) {
								employee.setEmpName(String.valueOf(cell.getNumericCellValue()));
							} else {
								employee.setEmpName(cell.getStringCellValue());
							}
							break;
						case 2:
							if (ReadExcelDemo.getCellType(cell) == 0) {
								employee.setEmpDate(String.valueOf(cell.getDateCellValue()));
							} else {
								employee.setEmpDate(String.valueOf(cell.getDateCellValue()));

							}
							if (isEmpIdMatches) {
								String month = ReadExcelDemo.getMonth(employee.getEmpDate());
								if (empMonth.equalsIgnoreCase(month)) {
									isEmpMonthMatches = true;
								}

							}

							break;
						case 3:
							if (ReadExcelDemo.getCellType(cell) == 0) {
								employee.setEmpCheckIn(String.valueOf(cell.getDateCellValue()));
							} else {
								employee.setEmpCheckIn(cell.getStringCellValue());

							}
							break;
						case 4:
							if (ReadExcelDemo.getCellType(cell) == 0) {
								employee.setEmpCheckOut(String.valueOf(cell.getDateCellValue()));
							} else {
								employee.setEmpCheckOut(cell.getStringCellValue());

							}
							if (isEmpMonthMatches) {
								String empCheckOutTime = ReadExcelDemo.getTime(employee.getEmpCheckOut());
								String timeTokens[] = empCheckOutTime.split(":");
								int Hours = Integer.valueOf(timeTokens[0]);
								if (Hours > 21) {
									nightShifts++;
								} else {
									dayShifts++;
								}
							}
							break;
						}
					}

				}
				if (i != 0)
					employees.add(employee);

			}

			for (Employee emp : employees) {
				System.out.println(emp.getEmpid() + "\t" + emp.getEmpName() + emp.getEmpDate() + "\t" + "\t"
						+ emp.getEmpCheckIn() + emp.getEmpCheckOut());
			}

			System.out.println("DayShifts=" + dayShifts + "\t\t" + "nightShifts =" + nightShifts);
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
