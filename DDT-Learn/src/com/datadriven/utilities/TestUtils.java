package com.datadriven.utilities;

import java.io.IOException;
import java.util.ArrayList;

import com.apache.poi.excel.XLSReader;

public class TestUtils {

	public static ArrayList<Object[]> getDataFormExcel(String fileExcelPath) {
		
		ArrayList<Object[]> allExcelData = new ArrayList<>();
		XLSReader xlsReader = null;
		try {
			xlsReader = new XLSReader(fileExcelPath);
			for (int i = 2; i <= xlsReader.getRowCount("RegTestData"); i++) {
				allExcelData.add(new Object[] { xlsReader.getCellData("RegTestData", "firstname", i),
						xlsReader.getCellData("RegTestData", "lastname", i),
						xlsReader.getCellData("RegTestData", "emailaddress", i),
						xlsReader.getCellData("RegTestData", "password", i) });
			}

			xlsReader.close();
		} catch (IOException e) {
			e.printStackTrace();
			return null;
		}

		return allExcelData;
	}

}
