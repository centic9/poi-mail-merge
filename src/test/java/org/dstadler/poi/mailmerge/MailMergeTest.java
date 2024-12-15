package org.dstadler.poi.mailmerge;

import java.io.File;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;

class MailMergeTest {
    private static final File RESULT_FILE = new File("build/Result.docx");

	@BeforeAll
	static void setUpClass() {
		assertTrue(new File("build").exists() || new File("build").mkdirs(), "Failed to create directory 'build'");
	}

	@BeforeEach
	void setUp() {
		// ensure the result file is not there
		assertTrue(!RESULT_FILE.exists() || RESULT_FILE.delete(), "File should not exist or we should be able to delete it, exist: " + RESULT_FILE.exists());
	}

    @Test
    void test() throws Exception {
        // use sample files to run a full merge
        MailMerge.main(new String[] {"samples/Template.docx", "samples/Lines.xlsx", RESULT_FILE.getPath()});

        // ensure the result file is written now
        assertTrue(RESULT_FILE.exists());
    }

    @Test
    void testCSV() throws Exception {
        // use sample files to run a full merge
        MailMerge.main(new String[] {"samples/Template.docx", "samples/Lines.csv", RESULT_FILE.getPath()});

        // ensure the result file is written now
        assertTrue(RESULT_FILE.exists());
    }

    @Test
    void testNoArgs() throws Exception {
        try {
            MailMerge.main(new String[] {});
            fail();
        } catch (@SuppressWarnings("unused") IllegalArgumentException e) {
            // expected here
        }
    }

    @Test
    void testMissingDoc() throws Exception {
        try {
            MailMerge.main(new String[] {"samples/Missing.docx", "samples/Lines.xlsx", RESULT_FILE.getPath()});
            fail();
        } catch (@SuppressWarnings("unused") IllegalArgumentException e) {
            // expected here
        }
    }

    @Test
    void testInvalidDoc() throws Exception {
        try {
            MailMerge.main(new String[] {"samples", "samples/Lines.xlsx", RESULT_FILE.getPath()});
            fail();
        } catch (@SuppressWarnings("unused") IllegalArgumentException e) {
            // expected here
        }
    }

    @Test
    void testMissingXlsx() throws Exception {
        try {
            MailMerge.main(new String[] {"samples/Template.docx", "samples/Missing.xlsx", RESULT_FILE.getPath()});
            fail();
        } catch (@SuppressWarnings("unused") IllegalArgumentException e) {
            // expected here
        }
    }

    @Test
    void testInvalidXlsx() throws Exception {
        try {
            MailMerge.main(new String[] {"samples/Template.docx", "samples", RESULT_FILE.getPath()});
            fail();
        } catch (@SuppressWarnings("unused") IllegalArgumentException e) {
            // expected here
        }
    }

    @Test
    void testTagSplitByFormatting() throws Exception {
        // use sample files to run a full merge
        MailMerge.main(new String[] {"samples/Template-TagSplitByFormatting.docx", "samples/Lines.xlsx", RESULT_FILE.getPath()});

        // ensure the result file is written now
        assertTrue(RESULT_FILE.exists());
    }

    @Test
    void testWithIncludeIndicator() throws Exception {
        System.setProperty("org.dstadler.poi.mailmerge.includeindicator", "Include");
        try {
            // use sample files to run a full merge
            MailMerge.main(new String[]{"samples/Template.docx", "samples/Lines.xlsx", RESULT_FILE.getPath()});

            // ensure the result file is written now
            assertTrue(RESULT_FILE.exists());
        } finally {
            System.clearProperty("org.dstadler.poi.mailmerge.includeindicator");
        }
    }

    @Test
    void testWithIncludeIndicatorNoSuchColumn() throws Exception {
        System.setProperty("org.dstadler.poi.mailmerge.includeindicator", "Include");
        try {
            try {
                MailMerge.main(new String[]{"samples/Template.docx", "samples/Lines.csv", RESULT_FILE.getPath()});
                fail("Should fail because the system property points to an non-existing column");
            } catch (IllegalStateException e) {
                // expected here
            }

            // ensure the result file is not written now
            assertFalse(RESULT_FILE.exists());
        } finally {
            System.clearProperty("org.dstadler.poi.mailmerge.includeindicator");
        }
    }
}
