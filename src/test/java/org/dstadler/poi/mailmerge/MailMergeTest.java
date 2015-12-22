package org.dstadler.poi.mailmerge;

import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.File;

import org.junit.Test;

public class MailMergeTest {
	private static final File RESULT_FILE = new File("build/Result.docx");

	@Test
	public void test() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		// ensure the result file is not there
		assertTrue(RESULT_FILE.exists() || RESULT_FILE.delete());

		// use sample files to run a full merge
		MailMerge.main(new String[] {"samples/Template.docx", "samples/Lines.xlsx", RESULT_FILE.getPath()});

		// ensure the result file is written now
		assertTrue(RESULT_FILE.exists());
	}

	@Test
	public void testCSV() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());


		// ensure the result file is not there
		assertTrue(RESULT_FILE.exists() || RESULT_FILE.delete());

		// use sample files to run a full merge
		MailMerge.main(new String[] {"samples/Template.docx", "samples/Lines.csv", RESULT_FILE.getPath()});

		// ensure the result file is written now
		assertTrue(RESULT_FILE.exists());
	}

	@Test
	public void testNoArgs() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		try {
			MailMerge.main(new String[] {});
			fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
	public void testMissingDoc() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		try {
			MailMerge.main(new String[] {"samples/Missing.docx", "samples/Lines.xlsx", RESULT_FILE.getPath()});
			fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
	public void testInvalidDoc() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		try {
			MailMerge.main(new String[] {"samples", "samples/Lines.xlsx", RESULT_FILE.getPath()});
			fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
	public void testMissingXlsx() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		try {
			MailMerge.main(new String[] {"samples/Template.docx", "samples/Missing.xlsx", RESULT_FILE.getPath()});
			fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}

	@Test
	public void testInvalidXlsx() throws Exception {
		assertTrue("Failed to create directory 'build'", new File("build").exists() || new File("build").mkdirs());

		try {
			MailMerge.main(new String[] {"samples/Template.docx", "samples", RESULT_FILE.getPath()});
			fail();
		} catch (@SuppressWarnings("unused") IllegalArgumentException e) {
			// expected here
		}
	}
}
