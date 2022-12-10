# Spreadsheet-Complexity-Analyser
The Spreadsheet Complexity Analyser is a prototype for the Archive Interest Group (AIG) of the Open Preservation Foundation (OPF, http://openpreservation.org/).
## Usage
Execute

java -jar SpreadsheetComplexityAnalyser.jar DIR

to process *.xl[st][xm] and *.xl[akms] files in DIR (DIR must be a directory).

Use the command line parameter -v for verbose text output. Use the command line parameter -x for verbose XML output (which suppresses -v). Use the command line parameter -r to also recurse into any subdirectories. The parameter -h will output help information.

Execute

java -jar SpreadsheetComplexityAnalyser.jar

without any paramters to get usage information.

SCA currently extracts:
- File: file size, creation date/time, last accessed, last modified
- Workbook: #worksheets, #fonts, #defined names, #cell styles, #external links, VBA macros (present or not, tentative) and revision history (credits: Rauno Umborg) 
- Per work sheet: #formulas, #hyperlinks, #cellComments, #shapes, #dates, #cells used, #physical cells used, #rows used, #physical rows used, #tables, #pivot tables and #charts.


## Motivation
The Open Preservation Foundation Archives Interest Group investigated the significant properties of spreadsheets. We wanted to find the best suited (spreadsheet) file format for preserving (significant properties of) spreadsheets. As part of this study, we wanted to be able to distinguish between 'simple/static' spreadsheets and 'complex/dynamic' spreadsheets.

We defined 'simple' spreadsheet as those that 'just' contain some rows and columns with cells with static values, and e.g. some formatting for pretty-printing. 'Complex' spreadsheet examples are those with macros, formulas, cells referring to cells (on other sheets), named objects, etc. (Yes, we there would be a lot of grey area spreadsheets.)

The main reason for making this distinction was (cost-efficiency w.r.t.) normalisation and preservation: our hypothesis was that simple spreadsheets can be normalised to and preserved as e.g. PDF(/A), while more complex spreadsheets require a spreadsheet-specific file format. There is a lot of knowledge of and expertise in working with PDF(/A) in archives. Being able to preserve some percentage of spreadsheets as PDF(/A) might be cost-efficient. Our work was also meant to result in choosing the best suited spreadsheet-specific file format for the complex spreadsheets.

### Please note: this hypothesis was rejected. Please preserve spreadsheets in spreadsheet formats. Too many properties are lost when you normalise/convert to non-spreadsheet specific file formats. Our investigation helped confirm this, and will impact on e.g. Danish and Dutch government preferred file formats norms. See our final report for more information: https://zenodo.org/record/5468116.

But anyway, in order to distinguish between these two types of spreadsheets, we needed to be able to extract information about cells, sheets, formulas, named objects, macros, etc. Property extraction tools like JHOVE, FITS and Apache Tika don't extract that kind of information, and that led to the development of the Spreadsheet Complexity Analyser.
## Technology used
The Spreadsheet Complexity Analyser is written in Java, and uses Apache POI-HSSF and POI-XSSF to access the Microsoft Excel spreadsheet formats (xls and xlsx) - which form the bulk of the spreadsheets that we (archives) receive.
## Installation
The SpreadsheetComplexityAnalyser is available as an executable Java 8 jar file. Please also download the .cfg and .xsd for SCA configuration and XML output validation.

Execute e.g.

java -jar SpreadsheetComplexityAnalyser.jar to get the usage information.
## Contribute
Even though our project ended and the AIG is disbanded, we would greatly appreciate contributions to this initiative. You can help improve the code and give feedback on our approach. Contributions are not limited to OPF members or archives, in the same way that preservation, spreadsheets and significant properties are not issues limited to OPF members or archives.
## Credits
Thank you core AIG members: Kati Sein (NAE), Anders Bo Nielsen (DNA), Phillip Mike Toemmerholt (DNA), Frederik Holmelund Kjaerskov (DNA), Jacob Takema (KB/NANETH), Jonathan Tilbury (Preservica), Jack O'Sullivan (Preservica), Becky McGuinness (OPF) and Pepijn Lucker (NANETH).
And thank you Carl Wilson (OPF) for getting the code to Github and sharing development best practices.
## License
The Spreadsheet Complexity Analyser is the result my work as a Preservation Officer at the National Archives of the Netherlands. We, and the Dutch government in general, want to make as much information available in the public domain. As a result, my intellectual property in the Spreadsheet Complexity Analyser has a CC0 license (https://creativecommons.org/choose/zero/). Included libraries may have their own licenses.
