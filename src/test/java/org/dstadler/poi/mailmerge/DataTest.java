package org.dstadler.poi.mailmerge;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import static org.junit.Assert.*;

public class DataTest {
	private static final File EMPTY_XLSX = new File("build/Empty.xlsx");

	@Test
	public void testRead() throws Exception {
		Data data = new Data();
		data.read(new File("samples/Lines.xlsx"));

		assertEquals("Had: " + data.getHeaders().toString(),
				"[Name, Organisation, Address, Zip, City, Salutation, Include, , , , , , , , , , , null]", data.getHeaders().toString());
		assertEquals(3, data.getData().size());
	}

	@Test
	public void testReadCSV() throws Exception {
		Data data = new Data();
		data.read(new File("samples/Lines.csv"));

		assertEquals("Had: " + data.getHeaders().toString(),
				"[Name, Organisation, Address, Zip, City, Salutation, ]", data.getHeaders().toString());
		assertEquals(3, data.getData().size());
	}

	@Test
	public void testReadEmptyExcel() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		// prepare a file without sheet
		try (Workbook wb = new XSSFWorkbook()) {
			try (OutputStream stream = new FileOutputStream(EMPTY_XLSX)) {
				wb.write(stream);
			}
		}

		Data data = new Data();
		try {
			data.read(EMPTY_XLSX);
			fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
	public void testReadExcelNoRow() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		// prepare a file without sheet
		try (Workbook wb = new XSSFWorkbook()) {
			wb.createSheet("somename1");
			wb.createSheet("somename2");
			wb.removeSheetAt(0);
			try (OutputStream stream = new FileOutputStream(EMPTY_XLSX)) {
				wb.write(stream);
			}
		}

		Data data = new Data();
		try {
			data.read(EMPTY_XLSX);
			fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
	public void testReadExcelEmptyRowInBetween() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		// prepare a file without sheet
		try (Workbook wb = new XSSFWorkbook()) {
			Sheet sheet = wb.createSheet("somename");
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue("Header");

			row = sheet.createRow(2);
			cell = row.createCell(0);
			cell.setCellValue("Value");

			try (OutputStream stream = new FileOutputStream(EMPTY_XLSX)) {
				wb.write(stream);
			}
		}

		Data data = new Data();
		data.read(EMPTY_XLSX);

		assertEquals("Had: " + data.getHeaders().toString(),
				"[Header, null]", data.getHeaders().toString());
		assertEquals(1, data.getData().size());
		assertEquals("Had: " + data.getData().toString(),
				"[[Value, null]]", data.getData().toString());
	}

	@Test
	public void testReadExcelFileWithoutSheet() throws Exception {
		Workbook wb = new XSSFWorkbook();

		File file = File.createTempFile("MailMergeDataTest", ".xlsx");
		try {
			assertTrue(file.delete());
			try (OutputStream stream = new FileOutputStream(file)) {
				wb.write(stream);
			}
			assertTrue(file.exists());

			Data data = new Data();

			try {
				data.read(file);
				fail("Will fail without sheet in the workbook");
			} catch (IllegalArgumentException e) {
				// expected here
				assertNotNull(e);
			}
		} finally {
			assertTrue(!file.exists() || file.delete());
		}
	}
}
