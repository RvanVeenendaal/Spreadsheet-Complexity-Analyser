# Spreadsheet-Complexity-Analyser
The Spreadsheet Complexity Analyser is a prototype for the Archive Interest Group (AIG) of the Open Preservation Foundation (OPF, http://openpreservation.org/).
## Motivation
As AIG, we are studying the significant properties of spreadsheets. We want to find the best suited (spreadsheet) file format for preserving (significant properties of) spreadsheets. As part of this study, we wanted to be able to distinguish between 'simple/static' spreadsheets and 'complex/dynamic' spreadsheets.

Simple spreadsheet examples are those that just contain some rows and columns with cells with static values, and e.g. some formatting for pretty-printing. Complex spreadsheet examples are those with macros, formulas, cells referring to cells (on other sheets), named objects, etc. (Yes, we know there will be a lot of grey area spreadsheets.) 

The main reason for making this distinction is (cost-efficiency w.r.t.) normalisation and preservation: our hypothesis is that simple spreadsheets can be normalised to and preserved as e.g. PDF(/A), while more complex spreadsheets require a spreadsheet-specific file format. There is a lot of knowledge of and expertise in working with PDF(/A) in archives. Being able to preserve some percentage of spreadsheets as PDF(/A) is cost-efficient. Our work should also result in choosing the best suited spreadsheet-specific file format for the complex spreadsheets.

In order to distinguish between these types of spreadsheets, we need to be able to extract information about cells, sheets, formulas, named objects, macros, etc. Property extraction tools like FITS, JHOVE and Apache Tika don't extract that kind of information, and that led to the development of the Spreadsheet Complexity Analyser.
## Technology used
The Spreadsheet Complexity Analyser is written in Java, and uses Apache POI (HSSF and XSSF) to access the Microsoft Excel spreadsheet formats (xls and xlsx) - which form the bulk of the spreadsheets that we (archives) receive.
## Installation
The SpreadsheetComplexityAnalyser is available as a runnable Java 8 jar file. 
Execute 'java -jar SpreadsheetComplexityAnalyser.jar' to get the usage information.
## Contribute
We would greatly appreciate contributions to this initiative. You can help improve the code, give feedback on our approach or otherwise contribute to the AIG. Contributions are not limited to OPF members or archives, in the same way that preservation, spreadsheets and significant properties are not issues limited to OPF members or archives.
## Credits
Thank you core AIG members: Kati Sein (NAE), Anders Bo Nielsen (DNA), Phillip Mike Toemmerholt (DNA), Frederik Holmelund Kjaerskov (DNA), Jacob Takema (KB/NANETH), Jonathan Tilbury (Preservica), Jack O'Sullivan (Preservica), Becky McGuinness (OPF) and Pepijn Lucker (NANETH).
And thank you Carl Wilson (OPF) for getting the code to Github and sharing development best practices.
## License
The Spreadsheet Complexity Analyser is the result my work as a Preservation Officer at the National Archives of the Netherlands. We, and the Dutch government in general, want to make as much information available in the public domain. As a result, my intellectual property in the Spreadsheet Complexity Analyser has a CC0 license (https://creativecommons.org/choose/zero/). Included libraries may have their own licenses.
