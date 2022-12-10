// Apache POI-XSSF and POI-HSSF used to read (Excel) spreadsheets
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFChart;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
// Other helper imports
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.attribute.FileTime;
import java.util.Collection;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.ParseException;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.DirectoryFileFilter;
import org.apache.commons.io.filefilter.WildcardFileFilter;

public class ApachePOIExcelReader {
	// Configuration 
	private static boolean verbose = false;
    private static boolean xml_out = false;
    private static boolean recursive = false;
    private static boolean config = false;
    private static boolean help = false;
    // Container for properties
    private static SpreadsheetProperties sp;
    // Workbook level properties
    private static int iFontsThreshold = 1;
    private static int iDefinedNamesThreshold = 1;
    private static int iCellStylesThreshold = 1;
    private static int iVBAMacrosThreshold = 0;
    private static int iExternalLinksThreshold = 0;
    private static int iHasRevisionHistoryThreshold = 0;
    // Worksheet level properties
    private static int iPivotTablesThreshold = 0;
    private static int iTablesThreshold = 0;
    private static int iRowsUsedThreshold = 1000;
    private static int iPhysicalRowsUsedThreshold = 1000;
    private static int iChartsThreshold = 0;
    private static int iWorksheetsThreshold = 1;
    private static int iFormulasThreshold = 0;
    private static int iHyperlinksThreshold = 0;
    private static int iCellCommentsThreshold = 0;
    private static int iShapesThreshold = 0;
    private static int iDatesThreshold = 0;
    private static int iCellsUsedThreshold = 1000;
    private static int iPhysicalCellsUsedThreshold = 1000;
    
    public static boolean processFile(File file, boolean isXSSF) {
    	System.err.println("Processing: "+ file);
    	boolean skip = false;	// skip this file is it is an unsupported Excel version
    	try {
    		if(file.exists()) { 
    			Path path = Paths.get(file.getCanonicalPath());
    			BasicFileAttributes attributes = Files.readAttributes(path, BasicFileAttributes.class);
    			// Get file information
    			double sizeInKb = attributes.size() / 1024;
    			FileTime creationTime = attributes.creationTime();
    			FileTime accessTime = attributes.lastAccessTime();
    			FileTime modifyTime = attributes.lastModifiedTime();
    			sp.getFileProperties().setdFileSizeKb(sizeInKb);
    			sp.getFileProperties().setsLastModified(modifyTime.toString());
    	        sp.getFileProperties().setsLastAccessed(accessTime.toString());
    	        sp.getFileProperties().setsCreation(creationTime.toString());
    		}
    		FileInputStream excelFile = new FileInputStream(file);
    		// Get workbook level information 
	        Workbook workbook = WorkbookFactory.create(excelFile);
	        sp.getWorkbookProperties().setiDefinedNames(workbook.getNumberOfNames());
	        sp.getWorkbookProperties().setiWorkSheets(workbook.getNumberOfSheets());
	        sp.getWorkbookProperties().setiFonts(workbook.getNumberOfFonts());
	        sp.getWorkbookProperties().setiCellStyles(workbook.getNumCellStyles());
	        if (isXSSF) {
		        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
		        sp.getWorkbookProperties().setiExternalLinks(xssfWorkbook.getExternalLinksTable().size());
	        }
	        else {
	        	sp.getWorkbookProperties().setiExternalLinks(-1); // -1 signals not available (for hssf/xls)
	        }
	        Iterator<Sheet> sheetIterator = workbook.iterator();
	        while (sheetIterator.hasNext()) {
	        	Sheet currentSheet = sheetIterator.next();
	        	WorksheetProperties worksheetProperties = new WorksheetProperties();
	        	// Get (work)sheet level information
	        	if (isXSSF) {
	        		XSSFSheet sheet = (XSSFSheet) currentSheet;
	        		worksheetProperties.setsSheetName(sheet.getSheetName());
	        		worksheetProperties.setiPivotTables(sheet.getPivotTables().size());
	        		worksheetProperties.setiTables(sheet.getTables().size());
	        		XSSFDrawing drawing = sheet.getDrawingPatriarch();
	        		if (drawing != null) {
		        		worksheetProperties.setiShapes(drawing.getShapes().size());
		        		worksheetProperties.setiCharts(drawing.getCharts().size());
		        	}
	        	}
	        	else {
	        		HSSFSheet sheet = (HSSFSheet) currentSheet;
	        		worksheetProperties.setsSheetName(sheet.getSheetName());
	        		worksheetProperties.setiPivotTables(-1);		// HSSF does not support pivot tables (2022-11-5)
	        		worksheetProperties.setiTables(-1);				// HSSF does not (seem to) support tables (2022-11-5)
	        		worksheetProperties.setiCharts(HSSFChart.getSheetCharts(sheet).length);
        		HSSFPatriarch drawing = (HSSFPatriarch) sheet.getDrawingPatriarch();
	        		if (drawing != null) {
	        			// System.err.println("Number of children of sheet " + sheet.getSheetName() + ": " + drawing.countOfAllChildren());
		        		worksheetProperties.setiShapes(worksheetProperties.getiShapes() + drawing.countOfAllChildren());        		
		        	}
	        	}
	        	worksheetProperties.setiRowsUsed(worksheetProperties.getiRowsUsed() + currentSheet.getLastRowNum());
	        	worksheetProperties.setiPhysicallyUsedRows(worksheetProperties.getiPhysicallyUsedRows() + currentSheet.getPhysicalNumberOfRows());

	        	// There is no interface for get(PhysicalNumberOf)Columns 
	        	// (but you could use the number of (physical) cells per row)

	            Iterator<Row> rowIterator = currentSheet.iterator();
	            while (rowIterator.hasNext()) {
	                Row currentRow = rowIterator.next();
	                worksheetProperties.setiCellsUsed(worksheetProperties.getiCellsUsed() + currentRow.getLastCellNum());
	                worksheetProperties.setiPhysicallyUsedCells(worksheetProperties.getiPhysicallyUsedCells() + currentRow.getPhysicalNumberOfCells());
		        	Iterator<Cell> cellIterator = currentRow.iterator();
	                while (cellIterator.hasNext()) {
	                    Cell currentCell = cellIterator.next();
	        			//System.err.println("Cell at row " + currentCell.getRowIndex() + " and column " + currentCell.getColumnIndex() + ": " +currentCell.getCellType());
	                    if (currentCell.getCellType() == CellType.FORMULA) {
	                    	worksheetProperties.setiFormulas(worksheetProperties.getiFormulas() + 1);
	                    }
	                    if (currentCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(currentCell)) {
	                    	worksheetProperties.setiDates(worksheetProperties.getiDates() + 1);
	                    }
	                    if(currentCell.getHyperlink() != null) {
	                    	worksheetProperties.setiHyperlinks(worksheetProperties.getiHyperlinks() + 1);                        	
	                    }
	                    if (currentCell.getCellComment() != null) {
	                    	worksheetProperties.setiCellComments(worksheetProperties.getiCellComments() + 1);
	                    }
	                }
	            }
	        	sp.getWorksheetPropertiesList().add(worksheetProperties);
	        }
	        workbook.close();
	        excelFile.close();
    	} catch (org.apache.poi.UnsupportedFileFormatException e) {
    		System.out.println("Skipping " + file.getName() + ". This file's format version is not supported.");
    		skip = true;
	    } catch (FileNotFoundException e) {
	    	System.out.println("Skipping " + file.getName() + ". Is this an Excel lock file?");
	    	skip = true;
	    } catch (IOException e) {
	    	System.out.println("Skipping " + file.getName() + ". Error reading (this type of Excel) file.");
	    	skip = true;
	    } catch (Exception e) {
	    	System.out.println("Skipping " + file.getName() + ". Exception");
	    	e.printStackTrace();
	    	skip = true;
	    }
    	if (!skip) {
    		if (isXSSF) {
    			findRevisionHeaders(file); // Only works for XLSX-family
    		}
    		else {
    			sp.getWorkbookProperties().setiHasRevisionHistory(-1); // Not supported for XLSX-family
    		}
    		findVBAMacros(file);
    	}
    	return skip;
	}

// To do: read threshold values for result calculation from config.ini file (or as CLI parameters?)    
// and use those values for calculating and outputting result simple/static or complex/dynamic
// but only when user asks for result calculation via CLI parameter
    public static void outputResults(File file) {   
		String result = "simple/static";
		// If ANY workbook value exceeds the threshold, it is considered complex/dynamic
		if (sp.getWorkbookProperties().getiWorkSheets() > iWorksheetsThreshold ||
	    		sp.getWorkbookProperties().getiFonts() > iFontsThreshold ||
	    		sp.getWorkbookProperties().getiDefinedNames() > iDefinedNamesThreshold ||
	    		sp.getWorkbookProperties().getiCellStyles() > iCellStylesThreshold ||
	    		sp.getWorkbookProperties().getiVBAMacros() > iVBAMacrosThreshold ||
	    		sp.getWorkbookProperties().getiExternalLinks() > iExternalLinksThreshold ||
	    		sp.getWorkbookProperties().getiHasRevisionHistory() > iHasRevisionHistoryThreshold
	    ) {
			result = "complex/dynamic";			
		}
		// If ANY worksheet has a value that exceeds the threshold, it is considered complex/dynamic
		for (WorksheetProperties worksheetProperties : sp.getWorksheetPropertiesList()) {
			if (worksheetProperties.getiFormulas() > iFormulasThreshold ||
					worksheetProperties.getiHyperlinks() > iHyperlinksThreshold ||
					worksheetProperties.getiCellComments() > iCellCommentsThreshold ||
					worksheetProperties.getiShapes() > iShapesThreshold ||
					worksheetProperties.getiDates() > iDatesThreshold ||
					worksheetProperties.getiCellsUsed() > iCellsUsedThreshold ||
					worksheetProperties.getiPhysicallyUsedCells() > iPhysicalCellsUsedThreshold ||
					worksheetProperties.getiPivotTables() > iPivotTablesThreshold ||
					worksheetProperties.getiTables() > iTablesThreshold ||
					worksheetProperties.getiCharts() > iChartsThreshold ||
					worksheetProperties.getiRowsUsed() > iRowsUsedThreshold ||
					worksheetProperties.getiPhysicallyUsedRows() > iPhysicalRowsUsedThreshold
			) {
				result = "complex/dynamic";				
			}
		}
    	if (xml_out) {
	        System.out.println("\t<spreadsheetComplexityAnalyserResult>");
	        System.out.println("\t\t<file name=\"" + file.getAbsoluteFile() + "\">");
	        System.out.println("\t\t\t<fileSize>" + sp.getFileProperties().getdFileSizeKb() + " kB</fileSize>");
	        System.out.println("\t\t\t<created>" + sp.getFileProperties().getsCreation() + "</created>");
	        System.out.println("\t\t\t<lastAccessed>" + sp.getFileProperties().getsLastAccessed() + "</lastAccessed>");	        
	        System.out.println("\t\t\t<lastModified>" + sp.getFileProperties().getsLastModified() + "</lastModified>");
	        System.out.println("\t\t</file>");
	        System.out.println("\t\t<workbook>");
	        System.out.println("\t\t\t<worksheets>" + sp.getWorksheetPropertiesList().size() + "</worksheets>");
	        System.out.println("\t\t\t<fonts>" + sp.getWorkbookProperties().getiFonts() + "</fonts>");
	        System.out.println("\t\t\t<definedNames>" + sp.getWorkbookProperties().getiDefinedNames() + "</definedNames>");
	        System.out.println("\t\t\t<cellStyles>" + sp.getWorkbookProperties().getiCellStyles() + "</cellStyles>");
	        System.out.println("\t\t\t<externalLinks>" + sp.getWorkbookProperties().getiExternalLinks() + "</externalLinks>");
	        System.out.println("\t\t\t<revisionHistory>" + sp.getWorkbookProperties().getiHasRevisionHistory() + "</revisionHistory>");
	        System.out.println("\t\t\t<vbaMacros>" + sp.getWorkbookProperties().getiVBAMacros() + "</vbaMacros>");
	        System.out.println("\t\t\t<worksheets>");
    	}
    	else if (verbose) {
    		System.out.println("spreadsheet:");
    		System.out.println("\tfile:");
    		System.out.println("\t\tname:\t\t\t" + file.getAbsoluteFile()); 
    		System.out.println("\t\tsize:\t\t\t" + sp.getFileProperties().getdFileSizeKb() + " kB"); 
    		System.out.println("\t\tcreated:\t\t" + sp.getFileProperties().getsCreation());
    		System.out.println("\t\tlast accessed:\t\t" + sp.getFileProperties().getsLastAccessed());
    		System.out.println("\t\tlast modified:\t\t" + sp.getFileProperties().getsLastModified());
    		System.out.println("\tworkbook:");
    		System.out.println("\t\tworksheets:\t\t" + sp.getWorkbookProperties().getiWorkSheets());
	        System.out.println("\t\tfonts:\t\t\t" + sp.getWorkbookProperties().getiFonts());
	        System.out.println("\t\tdefined names:\t\t" + sp.getWorkbookProperties().getiDefinedNames());
	        System.out.println("\t\tcell styles:\t\t" + sp.getWorkbookProperties().getiCellStyles());
	        System.out.println("\t\texternal links:\t\t" + sp.getWorkbookProperties().getiExternalLinks());
	        System.out.println("\t\trevision history:\t" + sp.getWorkbookProperties().getiHasRevisionHistory());
	        System.out.println("\t\tvba macros:\t\t" + sp.getWorkbookProperties().getiVBAMacros());
		}
		for (WorksheetProperties worksheetProperties : sp.getWorksheetPropertiesList()) {
			if (xml_out) {
		        System.out.println("\t\t\t\t<worksheet name=\"" + worksheetProperties.getsSheetName() + "\">");
   		        System.out.println("\t\t\t\t\t<formulas>" + worksheetProperties.getiFormulas() + "</formulas>");
		        System.out.println("\t\t\t\t\t<hyperlinks>" + worksheetProperties.getiHyperlinks() + "</hyperlinks>");
		        System.out.println("\t\t\t\t\t<cellComments>" + worksheetProperties.getiCellComments() + "</cellComments>");
		        System.out.println("\t\t\t\t\t<shapes>" + worksheetProperties.getiShapes() + "</shapes>");
		        System.out.println("\t\t\t\t\t<charts>" + worksheetProperties.getiCharts() + "</charts>");
		        System.out.println("\t\t\t\t\t<pivotTables>" + worksheetProperties.getiPivotTables() + "</pivotTables>");
		        System.out.println("\t\t\t\t\t<tables>" + worksheetProperties.getiTables() + "</tables>");
		        System.out.println("\t\t\t\t\t<dates>" + worksheetProperties.getiDates() + "</dates>");
		        System.out.println("\t\t\t\t\t<usedCells>" + worksheetProperties.getiCellsUsed() + "</usedCells>");
		        System.out.println("\t\t\t\t\t<physicallyUsedCells>" + worksheetProperties.getiPhysicallyUsedCells() + "</physicallyUsedCells>");
		        System.out.println("\t\t\t\t\t<usedRows>" + worksheetProperties.getiRowsUsed() + "</usedRows>");
		        System.out.println("\t\t\t\t\t<physicallyUsedRows>" + worksheetProperties.getiPhysicallyUsedRows() + "</physicallyUsedRows>");
		        System.out.println("\t\t\t\t</worksheet>");
	        }
			else if (verbose) {
				System.out.println("\t\tworksheet:");
				System.out.println("\t\t\tname:\t\t\t" + worksheetProperties.getsSheetName());
		        System.out.println("\t\t\tformulas:\t\t" + worksheetProperties.getiFormulas());
		        System.out.println("\t\t\thyperlinks:\t\t" + worksheetProperties.getiHyperlinks());
		        System.out.println("\t\t\tcellComments:\t\t" + worksheetProperties.getiCellComments());
		        System.out.println("\t\t\tshapes:\t\t\t" + worksheetProperties.getiShapes());
		        System.out.println("\t\t\tcharts:\t\t\t" + worksheetProperties.getiCharts());
		        System.out.println("\t\t\tpivotTables:\t\t" + worksheetProperties.getiPivotTables());
		        System.out.println("\t\t\ttables:\t\t\t" + worksheetProperties.getiTables());
		        System.out.println("\t\t\tdates:\t\t\t" + worksheetProperties.getiDates());
		        System.out.println("\t\t\tcells used:\t\t" + worksheetProperties.getiCellsUsed());
		        System.out.println("\t\t\tphysically used cells:\t" + worksheetProperties.getiPhysicallyUsedCells());
		        System.out.println("\t\t\trows used:\t\t" + worksheetProperties.getiRowsUsed());
		        System.out.println("\t\t\tphysically used rows:\t" + worksheetProperties.getiPhysicallyUsedRows());
			}
		}
	    if (xml_out) { 
	        System.out.println("\t\t\t</worksheets>");
	        System.out.println("\t\t</workbook>");
	        System.out.println("\t\t<tentativeAssessment>" + result + "</tentativeAssessment>");
	        System.out.println("\t</spreadsheetComplexityAnalyserResult>");
    	}
	    else if (verbose) {
	        System.out.println("\ttentative assessment:\t\t" + result + "\n");
	    }
    	else {
    	    System.out.println("Tentative spreadsheet complexity analyser result for\n\t" + file.getAbsoluteFile() + ": " + result + "\n");
    	}
    }
     
   public static void printHelpAndExit(String message, Options options) {
   		HelpFormatter formatter = new HelpFormatter(); 
	    formatter.printHelp("java -jar SpreadsheetComplexityAnalyser.jar DIR [-c] [-h] [-r] [-v] [-x]", options);
	    System.out.println(" DIR\t\t  directory with *.xl[st][xm] and *.xl[akms] files to process.");
	    System.out.println(message);
	    System.exit(0);
   }
   
   /*
    * Configuration file reader
    */
   public static void readConfigFile(String path) {
	   Properties prop = new Properties();
	   String fileName = path + "\\SpreadsheetComplexityAnalyser.cfg";
	   InputStream is = null;
	   try {
	       is = new FileInputStream(fileName);
	   } catch (FileNotFoundException ex) {
	       System.out.println("Error: config file not found! Using default values.");
	   }
	   try {
	       prop.load(is);
	   } catch (IOException ex) {
		   System.out.println("Error reading config file! Using default values.");
	   }
	   try {
		   // Try to read the values first...
		   iWorksheetsThreshold = Integer.parseInt(prop.getProperty("worksheetsThreshold"));
		   iFontsThreshold = Integer.parseInt(prop.getProperty("fontsThreshold"));
		   iDefinedNamesThreshold = Integer.parseInt(prop.getProperty("definedNamesThreshold"));
		   iCellStylesThreshold = Integer.parseInt(prop.getProperty("cellStylesThreshold"));
		   iFormulasThreshold = Integer.parseInt(prop.getProperty("formulasThreshold"));
		   iHyperlinksThreshold = Integer.parseInt(prop.getProperty("hyperlinksThreshold"));
		   iCellCommentsThreshold = Integer.parseInt(prop.getProperty("cellCommentsThreshold"));
		   iVBAMacrosThreshold = Integer.parseInt(prop.getProperty("vbaMacrosThreshold"));
		   iShapesThreshold = Integer.parseInt(prop.getProperty("shapesThreshold"));
		   iDatesThreshold = Integer.parseInt(prop.getProperty("datesThreshold"));
		   iCellsUsedThreshold = Integer.parseInt(prop.getProperty("cellsUsedThreshold"));
		   iPhysicalCellsUsedThreshold = Integer.parseInt(prop.getProperty("physicalCellsUsedThreshold"));
		   iRowsUsedThreshold = Integer.parseInt(prop.getProperty("rowsUsedThreshold"));
		   iPhysicalRowsUsedThreshold = Integer.parseInt(prop.getProperty("physicalRowsUsedThreshold"));
		   iExternalLinksThreshold = Integer.parseInt(prop.getProperty("externalLinksThreshold"));
		   iHasRevisionHistoryThreshold = Integer.parseInt(prop.getProperty("hasRevisionHistoryThreshold"));
		   iPivotTablesThreshold = Integer.parseInt(prop.getProperty("pivotTablesThreshold"));
		   iTablesThreshold = Integer.parseInt(prop.getProperty("tablesThreshold"));
		   iChartsThreshold = Integer.parseInt(prop.getProperty("chartsThreshold"));
	   }
	   catch (Exception e) {
		   System.out.println("Error reading config properties! Using default values.");
	   }
   }
   
/*
 * Spreadsheet Complexity Analyser
 */
    public static void main(String[] args) throws IOException {
    	CommandLineParser parser = new DefaultParser();
    	CommandLine cmd = null;
    	Options options = new Options();
    	options.addOption("v", "verbose", false, "verbose output: show number of occurrences of properties in text form" );
    	options.addOption("x", "xml", false, "xml output: show number of occurrences of properties in xml form (suppresses verbose output)");
    	options.addOption("r", "recursive", false, "recurse into subdirectories" );
    	options.addOption("c", "config", false, "config file: read complexity assessment threshold values from SpreadsheetComplexityAnalyser.cfg file");
    	options.addOption("h", "help", false, "help: show SpreadsheetComplexityAnalyser help information (and exit)");
    	try {
    	    // parse the command line arguments
    	    cmd = parser.parse( options, args );
    	}
    	catch( ParseException exp ) {
    	    printHelpAndExit("Error: cannot parse command line parameters:" + exp.getMessage(), options);
    	}
    	verbose = cmd.hasOption("verbose");
    	xml_out = cmd.hasOption("xml");
    	recursive = cmd.hasOption("recursive");
    	config = cmd.hasOption("config");
    	help = cmd.hasOption("help");

    	if (help) {
    		printHelpAndExit("\nHelp information for SpreadsheetComplexityAnalyser\n\n"
    				+ "This software extracts values of Excel spreadsheet properties and calculates\n"
    				+ "a tentative spreadsheet complexity assessment based on (default or config\n"
    				+ "file) threshold values.\n"
    				+ "Please note that the worksheet threshold values are used per worksheet:\n"
    				+ "\tit checks if any worksheet has a value that exceeds the threshold.\n\n"
    				+ "The assessment is 'simple/static' or 'complex/dynamic', but feel free to\n"
    				+ "ignore the assessment and use the extracted property values for other purposes.\n\n"
    				+ "This version can extract values for these properties:\n"
    				+ "file: file size, creation date/time, last accessed, last modified\n"
    				+ "workbook: worksheets, fonts, defined names, cell styles, external links, vba macros\n"
    				+ "\tand revision history\n"
    				+ "per sheet: formulas, hyperlinks, cellComments, shapes, dates, cells used, physical\n"
    				+ "\tcells used, rows used, physical rows used, tables, pivot tables and charts.\n"
    				+ "VBA macros: nonzero indicates possible VBA macros (tentative)\n\n"
    				+ "See the software's GitHub readme for more information:\n"
    				+ "https://github.com/RvanVeenendaal/Spreadsheet-Complexity-Analyser\n", options);
    	}
    	
    	if (cmd.getArgs().length < 1 || cmd.getArgs().length > 1) {
    		printHelpAndExit("Error: please provide exactly one input DIRectory.", options);
    	}
    	if (config) {
        	String path = new File(".").getCanonicalPath();
    		readConfigFile(path);
    	}
        File dir = new File(cmd.getArgs()[0]);
        if (!(dir.exists() && dir.isDirectory())) {
    		printHelpAndExit("Error: DIR " + dir + " does not exist or is not a directory.", options);
        }
    	WildcardFileFilter fileFilter = new WildcardFileFilter("*.*");
    	Collection<File> files = null;
    	if (recursive) {
    		files = FileUtils.listFiles(dir, fileFilter, DirectoryFileFilter.DIRECTORY);        	
    	}
    	else {
    		files = FileUtils.listFiles(dir, fileFilter, null);
    	}
    	Iterator<File> fileIterator = files.iterator();
    	if (xml_out) {
    		System.out.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
    		System.out.println("<spreadsheetComplexityAnalyserResults xmlns='http://openpreservation.org/spreadsheetComplexityAnalyser' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='SpreadsheetComplexityAnalyser.xsd'>");
    	}
    	else if (verbose) {
    		System.out.println("Spreadsheet complexity analyser results:");
    	}
        while (fileIterator.hasNext()) {        	
        	boolean skipped = false;
        	sp = new SpreadsheetProperties();
            File file = fileIterator.next();
            if (file.getName().toLowerCase().matches("^(?!\\~\\$)(.*)\\.xl[st][xm]$")) { //xlsx xlsm xltx xltm, no ~$ lock files
            	try {
                	skipped = processFile(file, true);
                	if (!skipped) {
                		outputResults(file);
                	}            		
            	}
            	catch (Exception e) {
            		printHelpAndExit("Error processing file:" + e.getMessage(), options);	
            	}            	
            }
            else if (file.getName().toLowerCase().matches("^(?!\\~\\$)(.*)\\.xl[akms]$")) { // xla xlk xlm xls, no ~$ lock files
            	try {
                	skipped = processFile(file, false);
                	if (!skipped) {
                		outputResults(file);
                	}            		
            	}
            	catch (Exception e) {
            		printHelpAndExit("Error processing file:" + e.getMessage(), options);	
            	}
            }
            else {
            	System.err.println("Unsupported file type: " + file.getName());
            }
    	}
    	if (xml_out) {
	        System.out.println("<legend>Legend: -1 = not supported (e.g. external links extraction for XLS). 0 or more = number of occurrences. At macros and revision history, nonzero means they are present.");
	        System.out.println("For more information about the extracted properties, see the Apache POI-HSSF or POI-XSSF at https://poi.apache.org/components/spreadsheet/index.html.</legend>");
    		System.out.println("</spreadsheetComplexityAnalyserResults>");
    	}
    	else if (verbose) {
	        System.out.println("legend:");
	        System.out.println("\t-1 = not supported (e.g. external links extraction for XLS).");
	        System.out.println("\t0 or more = number of occurrences.");
	        System.out.println("\tAt macros and revision history, nonzero means they are present.");
	        System.out.println("\tFor more information about the extracted properties, see the Apache POI-HSSF or POI-XSSF at https://poi.apache.org/components/spreadsheet/index.html.");
    	}        
        System.exit(0);
    }
    
    /*
    * Checks if binary file contains file with path and name of
    * x1/revisions/revisionHeaders.xml
    * Input should be XLSX file
    * Author: Rauno Umborg (rauno.umborg@ra.ee)
    */
    private static void findRevisionHeaders(File f){
    	// Load as binary:
        byte[] bytes = new byte[0];
        try {
            bytes = Files.readAllBytes(f.toPath());
        } catch (IOException e) {
            e.printStackTrace();
            return;
        }
        // Convert to string using UTF-8
        String asText = new String(bytes, StandardCharsets.UTF_8);
        // Find "xl/revisions/revisionHeaders.xml"
        int t = asText.indexOf("xl/revisions/revisionHeaders.xml");
        // If file exists, remember it.
        if (t > 0) {
        	sp.getWorkbookProperties().setiHasRevisionHistory(1);
        }
    }

    private static void findVBAMacros(File file) {
		try {
			FileInputStream excelFile = new FileInputStream(file);
	        VBAMacroReader reader = null;
	        Map<String, String> macros = null;
	        reader = new VBAMacroReader(excelFile);
	        macros = reader.readMacros();
	        Iterator<Entry<String, String>> macroIterator = macros.entrySet().iterator();
	        while(macroIterator.hasNext()) {
	        	continueWhile:
	        	{
		        	Map.Entry<String, String> macroEntry = macroIterator.next();
		        	String macro = macroEntry.getValue();
		        	String lines[] = macro.split("[\\n\\r]+");
		        	for (String line: lines){
		        		// Count only macros that actually have code: lines not starting with metadata key 'Attribute' 
		        		if (!line.matches("Attribute.*")) {
		        			sp.getWorkbookProperties().setiVBAMacros(sp.getWorkbookProperties().getiVBAMacros() + 1);
		        			break continueWhile;
		        		}
		        	}
	        	}
	        }
	        reader.close();
		} 
		catch (FileNotFoundException e) {
		    e.printStackTrace();
		} 
		catch (IOException e) {
		    e.printStackTrace();
		} 
		catch (IllegalArgumentException e) {
			// no VBA project found
		} 
		catch (Exception e) {
			e.printStackTrace();
		}
    }    
}