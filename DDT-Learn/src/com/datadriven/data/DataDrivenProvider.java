package com.datadriven.data;

import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;

import org.testng.annotations.DataProvider;

import com.datadriven.utilities.TestUtils;

public class DataDrivenProvider {

	public static final String INPUT_DATA_USER_EXCEL = "inputsDataUserExcel";
	public static final String FILE_EXCEL_PATH = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";

	
	@DataProvider(name=INPUT_DATA_USER_EXCEL)
	public Iterator<Object[]> getExcelDataUser() {

		ArrayList<Object[]> testData = TestUtils.getDataFormExcel(FILE_EXCEL_PATH);
		return testData.iterator();
	}

}
