package org.dstadler.poi.mailmerge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.dstadler.commons.logging.jdk.LoggerFactory;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

public class MailMerge {
	private static final Logger log = LoggerFactory.make();

	private List<String> headers = new ArrayList<>();
	private List<List<String>> values = new ArrayList<>();
	
	public static void main(String[] args) throws Exception {
		if(args.length != 3) {
			throw new IllegalArgumentException("Usage: MailMerge <word-template> <excel-template> <output-file>");
		}

		File wordTemplate = new File(args[0]);
		File excelFile = new File(args[1]);
		String outputFile = args[2];
		
		if(!wordTemplate.exists() || !wordTemplate.isFile()) {
			throw new IllegalArgumentException("Could not read Microsoft Word template " + wordTemplate);
		}
		if(!excelFile.exists() || !excelFile.isFile()) {
			throw new IllegalArgumentException("Could not read Microsoft Excel file " + excelFile);
		}
		
		new MailMerge().merge(wordTemplate, excelFile, outputFile);
	}

	private void merge(File wordTemplate, File excelFile, String outputFile) throws Exception {
		log.info("Merging data from " + wordTemplate + " and " + excelFile + " into " + outputFile);

		// read the lines from the Excel file
		readExcelFile(excelFile);
		
		// now open the word file and apply the changes
	    replace(wordTemplate, outputFile);

//	    OutputStream out = new FileOutputStream(outputFile);
//	    try {
//	    	doc.write(out);
//	    } finally {
//	    	out.close();
//	    }
	    
	    log.info("Done");
	}
	
	private void readExcelFile(File excelFile) throws EncryptedDocumentException, InvalidFormatException, IOException {
		try (Workbook wb = create(excelFile, true)) {
			Sheet sheet = wb.getSheetAt(0);
			if(sheet == null) {
				throw new IllegalArgumentException("Provided Microsoft Excel file " + excelFile + " does not have any sheet");
			}
	
			final int start;
			{ // read headers
				Row row = sheet.getRow(0);
				if(row == null) {
					throw new IllegalArgumentException("Provided Microsoft Excel file " + excelFile + " does not have data in the first row in the first sheet, "
							+ "but we expect the header data to be located there");
				}
				
				start = row.getFirstCellNum();
				for(int cellnum = start;cellnum < row.getLastCellNum();cellnum++) {
					Cell cell = row.getCell(cellnum);
					if(cell == null) {
						// add null to the headers if there a empty columns in the sheet
						headers.add(null);
						log.info("Had empty header for column " + CellReference.convertNumToColString(cellnum));
					} else {
						String value = cell.toString();
						headers.add(value);
						log.info("Had header '" + value + "' for column " + CellReference.convertNumToColString(cellnum));
					}
				}
			}
	
			for(int rownum = 1; rownum < sheet.getLastRowNum();rownum++) {
				Row row = sheet.getRow(rownum);
				if(row == null) {
					// ignore missing rows
					continue;
				}
			
				List<String> data = new ArrayList<>();
				for(int colnum = start;colnum < headers.size();colnum++) {
					Cell cell = row.getCell(colnum);  
					if(cell == null) {
						// store null-data for cells 
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

	@SuppressWarnings("resource")
    public static Workbook create(File file, boolean readOnly) throws IOException, InvalidFormatException, EncryptedDocumentException {
        if (! file.exists()) {
            throw new FileNotFoundException(file.toString());
        }

        try {
            NPOIFSFileSystem fs = new NPOIFSFileSystem(file, readOnly);
            return WorkbookFactory.create(fs);
        } catch(OfficeXmlFileException e) {
            // opening as .xls failed => try opening as .xlsx
            OPCPackage pkg = OPCPackage.open(file, readOnly ? PackageAccess.READ : PackageAccess.READ_WRITE);
            try {
                return new XSSFWorkbook(pkg);
            } catch (IOException ioe) {
                // ensure that file handles are closed (use revert() to not re-write the file)
                pkg.revert();
                //pkg.close();
                
                // rethrow exception
                throw ioe;
            } catch (IllegalArgumentException ioe) {
                // ensure that file handles are closed (use revert() to not re-write the file) 
                pkg.revert();
                //pkg.close();
                
                // rethrow exception
                throw ioe;
            }
        }
    }

	private void replace(File wordTemplate, String outputFile) throws XmlException, IOException {
		try (InputStream is = new FileInputStream(wordTemplate)) {
			XWPFDocument doc = new XWPFDocument(is);
		    CTBody body = doc.getDocument().getBody();
	
		    XmlOptions optionsOuter = new XmlOptions();
		    optionsOuter.setSaveOuter();
	
		    // read the current full Body text
		    String srcString = body.xmlText();
		    
		    // TODO: breaks appendBody() below...
		    //removeBody(body);
		    
		    // apply the replacements
		    //int docNr = 1;
		    for(List<String> data : values) {
		    	String replaced = srcString;
		    	for(int fieldNr = 0;fieldNr < headers.size();fieldNr++) {
		    		String header = headers.get(fieldNr);
		    		String value = data.get(fieldNr);
					if(header == null || value == null) {
		    			// ignore missing columns
		    			continue;
		    		}
		    		
					replaced = replaced.replace("${" + header + "}", value);
		    	}
		    	
		    	// TODO: combine all results into one document instead of writing multiple documents
			    /*CTBody makeBody = CTBody.Factory.parse(replaced);
			    body.set(makeBody);
			    
			    File resultFile = getOutputFile(outputFile, docNr);
			    log.info("Writing " + resultFile + " for " + data);
				try (OutputStream out = new FileOutputStream(resultFile)) {
			    	doc.write(out);
			    }*/
			    
				appendBody(body, replaced);
				
			    //docNr++;
		    }
		    
		    log.info("Writing overall result to " + outputFile);
			try (OutputStream out = new FileOutputStream(outputFile)) {
		    	doc.write(out);
		    }
		}
	}
	
	private static void appendBody(CTBody src, String append) throws XmlException {
	    XmlOptions optionsOuter = new XmlOptions();
	    optionsOuter.setSaveOuter();
	    String srcString = src.xmlText();
	    String prefix = srcString.substring(0,srcString.indexOf(">")+1);
	    String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
	    String sufix = srcString.substring( srcString.lastIndexOf("<") );
	    String addPart = append.substring(append.indexOf(">") + 1, append.lastIndexOf("<"));
	    CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
	    src.set(makeBody);
	}

//	private static void removeBody(CTBody src, String append) throws XmlException {
//	    XmlOptions optionsOuter = new XmlOptions();
//	    optionsOuter.setSaveOuter();
//	    String srcString = src.xmlText();
//	    String prefix = srcString.substring(0,srcString.indexOf(">")+1);
//	    String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
//	    String sufix = srcString.substring( srcString.lastIndexOf("<") );
//	    String addPart = append.substring(append.indexOf(">") + 1, append.lastIndexOf("<"));
//	    CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
//	    src.set(makeBody);
//	}
	
//	private static File getOutputFile(String outputFile, int i) {
//		String dir = FilenameUtils.getFullPath(outputFile);
//		String base = FilenameUtils.getBaseName(outputFile);
//		String ext = FilenameUtils.getExtension(outputFile);
//		
//		return new File(dir, base + "." + i + "." + ext);
//	}
}
