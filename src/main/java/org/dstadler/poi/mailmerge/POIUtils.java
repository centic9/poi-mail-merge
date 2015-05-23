package org.dstadler.poi.mailmerge;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIUtils {
	/**
	 * Helper methods because the {@link WorkbookFactory} in POI currently
	 * does not allow to open a Excel file read-only. 
	 * 
	 * We can remove this method when POI 3.13 is available.
	 * 
	 * @param file
	 * @param readOnly
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws EncryptedDocumentException
	 */
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

}
