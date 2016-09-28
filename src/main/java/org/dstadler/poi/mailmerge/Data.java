package org.dstadler.poi.mailmerge;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
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
import org.dstadler.commons.logging.jdk.LoggerFactory;

public class Data {
    private static final Logger log = LoggerFactory.make();

    private List<String> headers = new ArrayList<>();
    private List<List<String>> values = new ArrayList<>();

    public void read(File dataFile) throws IOException, EncryptedDocumentException, InvalidFormatException {

        // read the lines from the data-file
        if(FilenameUtils.getExtension(dataFile.getName()).equalsIgnoreCase("csv")) {
            readCSVFile(dataFile);
        } else {
            readExcelFile(dataFile);
        }

        removeEmptyLines();
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
                        //noinspection deprecation
                        switch (cell.getCellType()) {
                            //noinspection deprecation
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


    public List<List<String>> getData() {
        return values;
    }


    public List<String> getHeaders() {
        return headers;
    }
}
