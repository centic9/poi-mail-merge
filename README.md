[![Build Status](https://github.com/centic9/poi-mail-merge/actions/workflows/gradle-build.yml/badge.svg)](https://github.com/centic9/poi-mail-merge/actions)
[![Gradle Status](https://gradleupdate.appspot.com/centic9/poi-mail-merge/status.svg?branch=master)](https://gradleupdate.appspot.com/centic9/poi-mail-merge/status)

This is a small application which allows to repeatedly replace markers
in a Microsoft Word document with items taken from a CSV/Microsoft Excel 
file. 

I started this project as I was quite disappointed with the functionality 
that LibreOffice offers, I especially wanted something that is 
repeatable/automatable and does not produce spurious strange results and 
also does not need re-configuration each time the mail-merge is (re-)run.

## How it works

All you need is a Word-Document in "docx" format (>= 2003) which acts 
as template and an Excel .xls/.xlsx or CSV file which contains one row for 
each time the template-document should be produced.

The word-document can contain template-markers (enclosed in ${...}) for 
things that should be replaced, e.g. "${first-name} ${last-name}".

The first row of the first sheet of the Excel/CSV file is read as a 
header-row which is used to match the template-names used in the 
Word-template.

Only the first sheet of Excel files are read.

The result is a single merged Word-document which contains a replaced 
copy of the template for each line in the Excel file.

## Use it

### Grab and build it

    git clone https://github.com/centic9/poi-mail-merge.git
    cd poi-mail-merge
    ./gradlew installDist

### Run it

    ./run.sh <word-template> <excel/csv-file> <output-file>

### Sample files

There are some sample files in the directory `samples`, you can run these 
as follows

    ./gradlew installDist
    build\install\poi-mail-merge\bin\poi-mail-merge.bat samples\Template.docx samples\Lines.xlsx build\Result.docx

on Unix you can use the following steps

    ./gradlew installDist
    ./run.sh samples/Template.docx samples/Lines.xlsx build/Result.docx
	
## Support this project

If you find this tool useful and would like to support work on it, you can [Sponsor the author](https://github.com/sponsors/centic9)

## Tips

### Convert to PDF

You can use the tool ```unoconv``` from OpenOffice/LibreOffice to further 
convert the resulting docx, e.g. to PDF:

    unoconv -vvv --timeout=60 --doctype=document --output=result.pdf result.docx

## Known issues

### Only XLS/XLSX and one CSV format supported

For XLS/XLSX files only the first sheet is read and
headers are expected to be in the first row with data starting
in the second row.

For CSV, currently only files which use comma as delimiter and double-quotes 
for quoting text are supported. Other formats require code-changes, but should 
be easy to do by adjusting the CSFFormat definition (this project uses 
[Apache Commons CSV](http://commons.apache.org/proper/commons-csv/) for CSV handling).

### Only DOCX template format supported
 
The older .doc format is not supported as template document because this project 
makes heavy use of the internal XML format of DOCX files.

### High memory usage for large resulting files

The resulting output file is fully held in memory, so a very large number of
merged documents may cause very high memory usage and/or out-of-memory errors.

A streaming writing is currently not easy to support, but it should be possible
to add a mode of operation which writes separate files for the merged documents 
to overcome this limitation if necessary. Pull-requests highly welcome!

### Word-Formatting can confuse the replacement

If there are multiple formattings applied to a strings that holds a template-pattern, 
(e.g. if you make only half of the template-variable bold), the resulting 
XML-representation of the document might be split into multiple XML-Tags 
and thus might prevent the replacement from happening. 

A workaround is to use the formatting tool in LibreOffice/OpenOffice to ensure 
that the replacement tags have only one formatting applied to them. 

See centic9/poi-mail-merge#6 for possible improvements.

## Change it

### Build it and run tests

    cd poi-mail-merge
    ./gradlew check jacocoTestReport

Resulting coverage report is at `build/reports/jacoco/test/html/index.html`

#### Licensing

* poi-mail-merge is licensed under the [BSD 2-Clause License].

[BSD 2-Clause License]: https://www.opensource.org/licenses/bsd-license.php
