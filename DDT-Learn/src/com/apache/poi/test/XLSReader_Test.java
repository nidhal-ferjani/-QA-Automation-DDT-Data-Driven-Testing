package com.apache.poi.test;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Paths;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.apache.poi.excel.XLSReader;



public class XLSReader_Test {
	
	
	
	@Test(expectedExceptions=FileNotFoundException.class, enabled=false)
	public void testXLSReaderFileXlsNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData1.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		xlsReader.close();
	}
	
	@Test(enabled=false)
	public void testgetRouwCountWhenSheetExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		int rowCountActual =  xlsReader.getRowCount("RegTestData");
		xlsReader.close();
		Assert.assertEquals(17, rowCountActual);
		
	}
	

	@Test(enabled=false)
	public void testgetRouwCountWhenSheetExistEmpty() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		int rowCountActual =  xlsReader.getRowCount("HomePage");
		xlsReader.close();
	    Assert.assertEquals(0, rowCountActual);
		
	}
	
	@Test(enabled=false)
	public void testgetRouwCountWhenSheetNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		int rowCountActual =  xlsReader.getRowCount("RegTestDataX");
		xlsReader.close();
	    Assert.assertEquals(-1, rowCountActual);
		
	}
	
	
	@Test(enabled=false)
	public void testGetCellDataWhenSheetNameNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		String cellData = xlsReader.getCellData("nidhal sheet", "name", 1);
		xlsReader.close();
		Assert.assertNull(cellData);
	}

	
	@Test(enabled=false)
	public void testGetCellDataWhenRowNumberNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		String cellData = xlsReader.getCellData("RegTestData", "lastName", 18);
		xlsReader.close();
		Assert.assertNull(cellData);
	}

	@Test(enabled=false)
	public void testGetCellDataWhenCellExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		String cellData = xlsReader.getCellData("RegTestData", "lastName", 2);
		xlsReader.close();
		Assert.assertTrue(cellData.equalsIgnoreCase("peter"));;
	}
	
	@Test(enabled=false)
	public void testGetCellDataWhenCellValueIsDate() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\testfile.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		String cellData = xlsReader.getCellData("shEet1", "date", 2);
		xlsReader.close();
		Assert.assertTrue(cellData.equalsIgnoreCase("12/11/2014"));;
	}
	
	@Test(enabled=false)
	public void testGetCellDataWhenCellValueIsBlank() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\testfile.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		String cellData = xlsReader.getCellData("sheet1", "lastName", 2);
		xlsReader.close();
		Assert.assertTrue(cellData.equalsIgnoreCase(""));;
	}
	
	@Test(enabled=false)
	public void testGetCellDataCellValueByNumColNumRow() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		String cellData = xlsReader.getCellData("RegTestData",2,2);
		xlsReader.close();
		//System.out.println(cellData);
	    Assert.assertTrue(cellData.equalsIgnoreCase("peter"));;
	}
	
	@Test(enabled=false)
	public void testGetColumnCount() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		int colCount= xlsReader.getColumnCount("RegTestData");
		xlsReader.close();
		System.out.println(colCount);
	    Assert.assertEquals(colCount, 8);
	}
	
	@Test(enabled=false)
	public void testGetCelRowNum() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		int colCount= xlsReader.getCellRowNum("RegTestData", "emailaddress","test44@gmail.com");
		xlsReader.close();
	    Assert.assertEquals(colCount, 17);
	}
	@Test(enabled=false)
	public void testGetCelRowNumWhenCelValueNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		int colCount= xlsReader.getCellRowNum("RegTestData", "emailaddress","test444@gmail.com");
		xlsReader.close();
	    Assert.assertEquals(colCount, -1);
	}
	@Test(enabled=false)
	public void testAddSheetAlreadyExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		boolean status= xlsReader.addSheet("HomePage");
		xlsReader.close();
	    Assert.assertFalse(status);
	}
	@Test(enabled=false)
	public void testAddSheetNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		boolean status= xlsReader.addSheet("HomePageTest");
		xlsReader.close();
	    Assert.assertTrue(status);
	}
	
	@Test(enabled=false)
	public void testaddColumnSheetNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		boolean status= xlsReader.addColumn("HomePageTest1111","LastName");
		xlsReader.close();
	    Assert.assertFalse(status);
	}
	
	@Test(enabled=false)
	public void testaddColumnExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		boolean status= xlsReader.addColumn("RegTestData","Brith Day");
		xlsReader.close();
	    Assert.assertTrue(status);
	}
	
	@Test(enabled=false)
	public void testremoveColumnSheetNotExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		boolean status= xlsReader.removeColumn("HomePageTest1111",4);
		xlsReader.close();
	    Assert.assertFalse(status);
	}
	
	@Test(enabled=false)
	public void testremoveColumnExist() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		boolean status= xlsReader.removeColumn("RegTestData",6);
		xlsReader.close();
	    Assert.assertTrue(status);
	}
	
	@Test(enabled=true)
	public void testsetCellData() throws FileNotFoundException, IOException {
		String path = Paths.get("").toAbsolutePath() + "\\testData\\HalfEbayTestData.xlsx";
		XLSReader xlsReader = new XLSReader(path);
		boolean status= xlsReader.setCellData("HomePage", "emailaddress", 2, "nidhal.ferjani@gmail.com");
		xlsReader.close();
	   Assert.assertTrue(status);
	}
}
