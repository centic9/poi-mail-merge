package org.dstadler.poi.mailmerge;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

class DataTest {
	private static final File EMPTY_XLSX = new File("build/Empty.xlsx");

	@Test
    void testRead() throws Exception {
		Data data = new Data();
		data.read(new File("samples/Lines.xlsx"));

		Assertions.assertEquals("[Name, Organisation, Address, Zip, City, Salutation, Include, Date, , , , , , , , , , null]", data.getHeaders().toString());
		Assertions.assertEquals(4, data.getData().size());
	}

	@Test
    void testReadCSV() throws Exception {
		Data data = new Data();
		data.read(new File("samples/Lines.csv"));

		Assertions.assertEquals("[Name, Organisation, Address, Zip, City, Salutation]", data.getHeaders().toString());
		Assertions.assertEquals(3, data.getData().size());
	}

	@Test
    void testReadEmptyExcel() throws Exception {
		Assertions.assertTrue(new File("build").exists() || new File("build").mkdirs(), "Failed to create directory 'build'");

		// prepare a file without sheet
		try (Workbook wb = new XSSFWorkbook()) {
			try (OutputStream stream = new FileOutputStream(EMPTY_XLSX)) {
				wb.write(stream);
			}
		}

		Data data = new Data();
		try {
			data.read(EMPTY_XLSX);
			Assertions.fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
    void testReadExcelNoRow() throws Exception {
		Assertions.assertTrue(new File("build").exists() || new File("build").mkdirs(), "Failed to create directory 'build'");

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
			Assertions.fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
    void testReadExcelEmptyRowInBetween() throws Exception {
		Assertions.assertTrue(new File("build").exists() || new File("build").mkdirs(), "Failed to create directory 'build'");

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

		Assertions.assertEquals("[Header, null]", data.getHeaders().toString(), "Had: " + data.getHeaders().toString());
		Assertions.assertEquals(1, data.getData().size());
		Assertions.assertEquals("[[Value, null]]", data.getData().toString(), "Had: " + data.getData().toString());
	}

	@Test
    void testReadExcelFileWithoutSheet() throws Exception {
		File file = File.createTempFile("MailMergeDataTest", ".xlsx");
		try {
			Assertions.assertTrue(file.delete());

			try (Workbook wb = new XSSFWorkbook()) {
				try (OutputStream stream = new FileOutputStream(file)) {
					wb.write(stream);
				}
				Assertions.assertTrue(file.exists());

				Data data = new Data();

				try {
					data.read(file);
					Assertions.fail("Will fail without sheet in the workbook");
				} catch (IllegalArgumentException e) {
					// expected here
					Assertions.assertNotNull(e);
				}
			}
		} finally {
			Assertions.assertTrue(!file.exists() || file.delete());
		}
	}
}
