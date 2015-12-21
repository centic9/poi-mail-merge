package org.dstadler.poi.mailmerge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.dstadler.commons.logging.jdk.LoggerFactory;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

/**
 * Simple application which performs a "mail-merge" of a Microsoft Word template
 * document which contains replacement templates in the form of ${name}, ${first-name}, ...
 * and an Microsoft Excel spreadsheet which contains a list of entries that are merged in.
 *
 * Call this application with parameters <word-template> <excel/csv-template> <output-file>
 *
 * The resulting document has all resulting documents concatenated.
 *
 * @author dominik.stadler
 *
 */
public class MailMerge {
	private static final Logger log = LoggerFactory.make();

	private List<String> headers = new ArrayList<>();
	private List<List<String>> values = new ArrayList<>();

	public static void main(String[] args) throws Exception {
		LoggerFactory.initLogging();

		if(args.length != 3) {
			throw new IllegalArgumentException("Usage: MailMerge <word-template> <excel/csv-template> <output-file>");
		}

		File wordTemplate = new File(args[0]);
		File excelFile = new File(args[1]);
		String outputFile = args[2];

		if(!wordTemplate.exists() || !wordTemplate.isFile()) {
			throw new IllegalArgumentException("Could not read Microsoft Word template " + wordTemplate);
		}
		if(!excelFile.exists() || !excelFile.isFile()) {
			throw new IllegalArgumentException("Could not read data file " + excelFile);
		}

		new MailMerge().merge(wordTemplate, excelFile, outputFile);
	}

	private void merge(File wordTemplate, File dataFile, String outputFile) throws Exception {
		log.info("Merging data from " + wordTemplate + " and " + dataFile + " into " + outputFile);

		// read the lines from the data-file
		if(FilenameUtils.getExtension(dataFile.getName()).equals("csv")) {
			readCSVFile(dataFile);
		} else {
			readExcelFile(dataFile);
		}

		removeEmptyLines();

		// now open the word file and apply the changes
		try (InputStream is = new FileInputStream(wordTemplate)) {
			try (XWPFDocument doc = new XWPFDocument(is)) {
				// apply the lines and concatenate the results into the document
				applyLines(doc, outputFile);

			    log.info("Writing overall result to " + outputFile);
				try (OutputStream out = new FileOutputStream(outputFile)) {
			    	doc.write(out);
			    }
			}
		}
	}

	private void removeEmptyLines() {
		Iterator<List<String>> it = values.iterator();
		while(it.hasNext()) {
			List<String> line = it.next();
			boolean empty = true;
			for(String item : line) {
				if(StringUtils.isNotBlank(item)) {
					empty = false;
					break;
				}
			}

			// remove empty line
			if(empty) {
				log.info("Removing an empty data line");
				it.remove();
			}
		}
	}

	private void readCSVFile(File csvFile) throws IOException {
		// open file
		// List<String> lines = FileUtils.readLines(file, null);
		try (Reader reader = new FileReader(csvFile)) {
			CSVFormat strategy = CSVFormat.DEFAULT.
					withHeader().
			        withDelimiter(',').
                    withQuote('"').
			        withCommentMarker((char)0).
			        withIgnoreEmptyLines().
			        withIgnoreSurroundingSpaces();

			try (CSVParser parser = new CSVParser(reader, strategy)) {
				Map<String, Integer> headerMap = parser.getHeaderMap();
				for(Map.Entry<String,Integer> entry : headerMap.entrySet()) {
					headers.add(entry.getKey());
					log.info("Had header '" + entry.getKey() + "' for column " + entry.getValue());
				}

				List<CSVRecord> lines = parser.getRecords();
                log.info("Found " + lines.size() + " lines");
                for(CSVRecord line : lines) {
	                List<String> data = new ArrayList<>();
	                for(int pos = 0;pos < headerMap.size();pos++) {
						if(line.size() <= pos) {
							data.add(null);
						} else {
							data.add(line.get(pos));
						}
					}

					values.add(data);
				}
			}
		}
	}

	private void readExcelFile(File excelFile) throws EncryptedDocumentException, InvalidFormatException, IOException {
		try (Workbook wb = WorkbookFactory.create(excelFile, null, true)) {
			Sheet sheet = wb.getSheetAt(0);
			if(sheet == null) {
				throw new IllegalArgumentException("Provided Microsoft Excel file " + excelFile + " does not have any sheet");
			}

			final int start;
			final int end;
			{ // read headers
				Row row = sheet.getRow(0);
				if(row == null) {
					throw new IllegalArgumentException("Provided Microsoft Excel file " + excelFile + " does not have data in the first row in the first sheet, "
							+ "but we expect the header data to be located there");
				}

				start = row.getFirstCellNum();
				end = row.getLastCellNum();
				for(int cellnum = start;cellnum <= end;cellnum++) {
					Cell cell = row.getCell(cellnum);
					if(cell == null) {
						// add null to the headers if there are columns without title in the sheet
						headers.add(null);
						log.info("Had empty header for column " + CellReference.convertNumToColString(cellnum));
					} else {
						String value = cell.toString();
						headers.add(value);
						log.info("Had header '" + value + "' for column " + CellReference.convertNumToColString(cellnum));
					}
				}
			}

			for(int rownum = 1; rownum <= sheet.getLastRowNum();rownum++) {
				Row row = sheet.getRow(rownum);
				if(row == null) {
					// ignore missing rows
					continue;
				}

				List<String> data = new ArrayList<>();
				for(int colnum = start;colnum <= end;colnum++) {
					Cell cell = row.getCell(colnum);
					if(cell == null) {
						// store null-data for empty/missing cells
						data.add(null);
					} else {
						final String value;
				        switch (cell.getCellType()) {
			            case Cell.CELL_TYPE_NUMERIC:
			            	// ensure that numeric are formatted the same way as in the Excel file.
			            	value = CellFormat.getInstance(cell.getCellStyle().getDataFormatString()).apply(cell).text;
			                break;
			            default:
			            	// all others can use the default value from toString() for now.
			            	value = cell.toString();
				        }

						data.add(value);
					}
				}

				values.add(data);
			}
		}
	}


	private void applyLines(XWPFDocument doc, String outputFile) throws XmlException, IOException {
	    CTBody body = doc.getDocument().getBody();

	    XmlOptions optionsOuter = new XmlOptions();
	    optionsOuter.setSaveOuter();

	    // read the current full Body text
	    String srcString = body.xmlText();

	    // apply the replacements
	    boolean first = true;
	    for(List<String> data : values) {
	    	log.info("Applying to template: " + data);

	    	String replaced = srcString;
	    	for(int fieldNr = 0;fieldNr < headers.size();fieldNr++) {
	    		String header = headers.get(fieldNr);
	    		String value = data.get(fieldNr);

	    		// ignore columns without headers as we cannot match them
				if(header == null) {
	    			continue;
	    		}

				// use empty string for data-cells that have no value
				if(value == null) {
					value = "";
				}

				replaced = replaced.replace("${" + header + "}", value);
	    	}

			appendBody(body, replaced, first);

			first = false;
	    }
	}

	private static void appendBody(CTBody src, String append, boolean first) throws XmlException {
	    XmlOptions optionsOuter = new XmlOptions();
	    optionsOuter.setSaveOuter();
	    String srcString = src.xmlText();
	    String prefix = srcString.substring(0,srcString.indexOf(">")+1);

	    final String mainPart;
	    // exclude template itself in first appending
	    if(first) {
	    	mainPart = "";
	    } else {
	    	mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
	    }

	    String sufix = srcString.substring( srcString.lastIndexOf("<") );
	    String addPart = append.substring(append.indexOf(">") + 1, append.lastIndexOf("<"));
	    CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
	    src.set(makeBody);
	}
}
