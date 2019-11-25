import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
    private static boolean verbose = false;
    private static boolean xml_out = false;
    private static boolean recursive = false;
    private static boolean config = false;
    private static boolean help = false;
    private static SpreadsheetProperties sp;
    private static int worksheetsThreshold = 1;
    private static int fontsThreshold = 1;
    private static int definedNamesThreshold = 1;
    private static int cellStylesThreshold = 1;
    private static int formulasThreshold = 0;
    private static int hyperlinksThreshold = 0;
    private static int commentsThreshold = 0;
    private static int vbaMacrosThreshold = 0;
    private static int shapesThreshold = 0;
    private static int datesThreshold = 0;
    private static int cellsUsedThreshold = 1000;
    private static int externalLinksThreshold = 0;

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
            return ;
        }
        // Convert to string using UTF-8
        String asText = new String(bytes, StandardCharsets.UTF_8);
        // Find "xl/revisions/revisionHeaders.xml"
        int t = asText.indexOf("xl/revisions/revisionHeaders.xml");
        // If file exists, remember it.
        if (t > 0) {
        	sp.setHasRevisionHistory(1);
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
		        			sp.setHasVBAMacros(sp.getHasVBAMacros() + 1);
		        			break continueWhile;
		        		}
		        	}
	        	}
	        }
	        /* macros contains modules that might be empty
	         * macros also contains Excel objects like Sheet1 and ThisWorkbook
	         * therefore, we cannot be sure that there actually is a vba macro
	        */
//	        if (macros.size() > 0) { // due to Sheet1 and ThisWorkbook, perhaps > 2 ?
//	            sp.setHasVBAMacros(1);
//	        }            
	        reader.close();
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		} catch (IllegalArgumentException e) {
			// no VBA project found
		} catch (Exception e) {
			e.printStackTrace();
		}
    }    
    
    public static boolean processFile(File file, boolean isXSSF) {
    	boolean skip = false;	// skip this file is it is an unsupported Excel version
    	try {
    		if(file.exists()) { 
    			Path path = Paths.get(file.getCanonicalPath());
    			BasicFileAttributes attributes = Files.readAttributes(path, BasicFileAttributes.class);
    			double sizeInKb = attributes.size() / 1024;
    			FileTime creationTime = attributes.creationTime();
    			FileTime accessTime = attributes.lastAccessTime();
    			FileTime modifyTime = attributes.lastModifiedTime();
    			sp.setFileSizeKb(sizeInKb);
    			sp.setLastModified(modifyTime.toString());
    	        sp.setLastAccessed(accessTime.toString());
    	        sp.setCreation(creationTime.toString());
    		}
    		FileInputStream excelFile = new FileInputStream(file);
	        Workbook workbook;
	        workbook = WorkbookFactory.create(excelFile);
	        sp.setDefinedNames(workbook.getNumberOfNames());
	        sp.setWorkSheets(workbook.getNumberOfSheets());
	        sp.setFonts(workbook.getNumberOfFonts());
	        sp.setCellStyles(workbook.getNumCellStyles());
	        if (isXSSF) {
		        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
		        sp.setExternalLinks(xssfWorkbook.getExternalLinksTable().size());
	        }
	        else {
	        	sp.setExternalLinks(-1); // not available for hssf/xls
	        }
	        Iterator<Sheet> sheetIterator = workbook.iterator();
	        while (sheetIterator.hasNext()) {
	        	Sheet currentSheet = sheetIterator.next();
	        	if (isXSSF) {
	        		XSSFSheet sheet = (XSSFSheet) currentSheet;
	        		XSSFDrawing drawing = sheet.getDrawingPatriarch();
	        		if (drawing != null) {
		        		Iterator<XSSFShape> shapeIterator = drawing.iterator();
		        		int i = 0;
		        		while (shapeIterator.hasNext()){
//		        			XSSFShape shape = shapeIterator.next();
		        			shapeIterator.next();
// uncomment for testing 	System.out.println(sheet.getSheetName() + ", " + shape.getShapeName() + ", " + shape.getClass());
		        			i++;
		        		}
		        		sp.setShapes(sp.getShapes() + i);	        		
		        	}
	        	}
	        	else {
	        		HSSFSheet sheet = (HSSFSheet) currentSheet;
	        		HSSFPatriarch drawing = (HSSFPatriarch) sheet.getDrawingPatriarch();
	        		if (drawing != null) {
		        		Iterator<HSSFShape> shapeIterator = drawing.iterator();
		        		int i = 0;
		        		while (shapeIterator.hasNext()){
//		        			HSSFShape shape = shapeIterator.next();
		        			shapeIterator.next();
// uncomment for testing	System.out.println(sheet.getSheetName() + ", " + shape.getShapeName() + ", " + shape.getClass());
		        			i++;
		        		}
		        		sp.setShapes(sp.getShapes() + i);	        		
		        	}
	        	}
	            Iterator<Row> rowIterator = currentSheet.iterator();
	            while (rowIterator.hasNext()) {
	                Row currentRow = rowIterator.next();
		        	sp.setCellsUsed(sp.getCellsUsed() + currentRow.getPhysicalNumberOfCells());
	                Iterator<Cell> cellIterator = currentRow.iterator();
	                while (cellIterator.hasNext()) {
	                    Cell currentCell = cellIterator.next();
	                    if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
	                        sp.setFormulas(sp.getFormulas() + 1);
	                    }
	                    if (currentCell.getCellTypeEnum() == CellType.NUMERIC && DateUtil.isCellDateFormatted(currentCell)) {
	                        sp.setDates(sp.getDates() + 1);
	                    }
	                    if(currentCell.getHyperlink() != null) {
	                    	sp.setHyperlinks(sp.getHyperlinks() + 1);                        	
	                    }
	                    if (currentCell.getCellComment() != null) {
	                    	sp.setComments(sp.getComments() + 1);
	                    }
	                }
	            }            	
	        }
	        workbook.close();
	        excelFile.close();
    	} catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
    		System.out.println("Skipping " + file.getName() + ". This file's format version is not supported.");
    		skip = true;
	    } catch (FileNotFoundException e) {
	    	System.out.println("Skipping " + file.getName() + ". Is this an Excel lock file?");
	    	skip = true;
	    } catch (IOException e) {
	        e.printStackTrace();
	    } catch (Exception e) {
	    	e.printStackTrace();
	    }
    	if (!skip) {
    		if (isXSSF) {
    			findRevisionHeaders(file); // Only works for XLSX-family
    		}
    		else {
    			sp.setHasRevisionHistory(-1); // Not supported for XLSX-family
    		}
    		findVBAMacros(file);
    	}
    	return skip;
	}

    
// To do: read threshold values for result calculation from config.ini file (or as CLI parameters?)    
// and use those values for calculating and outputting result simple/static or complex/dynamic
// but only when user asks for result calculation via CLI parameter
    public static void outputResults(File file) {   
		String result = null;	    
		if (sp.getWorkSheets() > worksheetsThreshold ||
	    		sp.getFonts() > fontsThreshold ||
	    		sp.getDefinedNames() > definedNamesThreshold ||
	    		sp.getCellStyles() > cellStylesThreshold ||
	    		sp.getFormulas() > formulasThreshold ||
	    		sp.getHyperlinks() > hyperlinksThreshold ||
	    		sp.getComments() > commentsThreshold ||
	    		sp.getHasVBAMacros() > vbaMacrosThreshold ||
	    		sp.getShapes() > shapesThreshold ||
	    		sp.getDates() > datesThreshold ||
	    		sp.getCellsUsed() > cellsUsedThreshold ||
				sp.getExternalLinks() > externalLinksThreshold) {
	    	result = "complex/dynamic";
	    }
	    else {
	    	result = "simple/static";
	    }    	
    	if (xml_out) {
	        System.out.println("<spreadsheetComplexityAnalyserResult>");
	        System.out.println("\t<file>" + file.getAbsoluteFile() + "</file>");
//	        System.out.println("\t<result>" + result + "</result>");
	        System.out.println("\t<fileSize>" + sp.getFileSizeKb() + " kB</fileSize>");
	        System.out.println("\t<created>" + sp.getCreation() + "</created>");
	        System.out.println("\t<lastAccessed>" + sp.getLastAccessed() + "</lastAccessed>");	        
	        System.out.println("\t<lastModified>" + sp.getLastModified() + "</lastModified>");
	        System.out.println("\t<worksheets>" + sp.getWorkSheets() + "</worksheets>");
	        System.out.println("\t<fonts>" + sp.getFonts() + "</fonts>");
	        System.out.println("\t<definedNames>" + sp.getDefinedNames() + "</definedNames>");
	        System.out.println("\t<cellStyles>" + sp.getCellStyles() + "</cellStyles>");
	        System.out.println("\t<formulas>" + sp.getFormulas() + "</formulas>");
	        System.out.println("\t<hyperlinks>" + sp.getHyperlinks() + "</hyperlinks>");
	        System.out.println("\t<comments>" + sp.getComments() + "</comments>");
	        System.out.println("\t<vbaMacros>" + sp.getHasVBAMacros() + "</vbaMacros>");
	        System.out.println("\t<shapes>" + sp.getShapes() + "</shapes>");
	        System.out.println("\t<dates>" + sp.getDates() + "</dates>");
	        System.out.println("\t<usedCells>" + sp.getCellsUsed() + "</usedCells>");
	        System.out.println("\t<externalLinks>" + sp.getExternalLinks() + "</externalLinks>");
	        System.out.println("\t<revisionHistory>" + sp.getHasRevisionHistory() + "</revisionHistory>");
	        System.out.println("\t<tentativeAssessment>" + result + "</tentativeAssessment>");
	        System.out.println("</spreadsheetComplexityAnalyserResult>");
    	}
    	else if (verbose) {
    		//    	    System.out.println("Result: " + result);
    		System.out.println("File: " + file.getAbsoluteFile());
    		System.out.println("\tsize:\t\t\t" + sp.getFileSizeKb() + " kB"); 
    		System.out.println("\tcreated:\t\t" + sp.getCreation());
    		System.out.println("\tlast accessed:\t\t" + sp.getLastAccessed());
    		System.out.println("\tlast modified:\t\t" + sp.getLastModified());
	        System.out.println("\tworksheets:\t\t" + sp.getWorkSheets());
	        System.out.println("\tfonts:\t\t\t" + sp.getFonts());
	        System.out.println("\tdefined names:\t\t" + sp.getDefinedNames());
	        System.out.println("\tcell styles:\t\t" + sp.getCellStyles()); 
	        System.out.println("\tformulas:\t\t" + sp.getFormulas());
	        System.out.println("\thyperlinks:\t\t" + sp.getHyperlinks());
	        System.out.println("\tcomments:\t\t" + sp.getComments());
	        System.out.println("\tvba macros:\t\t" + sp.getHasVBAMacros());
	        System.out.println("\tshapes:\t\t\t" + sp.getShapes());
	        System.out.println("\tdates:\t\t\t" + sp.getDates());
	        System.out.println("\tcells used:\t\t" + sp.getCellsUsed());
	        System.out.println("\texternal links:\t\t" + sp.getExternalLinks());
	        System.out.println("\trevision history:\t" + sp.getHasRevisionHistory());
	        System.out.println("\ttentative assessment:\t" + result);
		}
    	else {
    	    System.out.println("Tentative spreadsheet complexity analyser result for\n\t" + file.getAbsoluteFile() + ": " + result);
    	}
    }
     
   public static void printHelpAndExit(String message, Options options) {
   		HelpFormatter formatter = new HelpFormatter(); 
   		System.out.println(message);
	    formatter.printHelp("java -jar SpreadsheetComplexityAnalyser.jar DIR [-c] [-h] [-r] [-v] [-x]", options);
	    System.out.println(" DIR\t\t  directory with *.xl[st][xm] and *.xl[akms] files to process.");
	    System.exit(0);
   }
   
   /*
    * Configuration file reader
    */
   public static void readConfigFile(String path) {
	   Properties prop = new Properties();
	   String fileName = path + "\\SpreadsheetComplexityAnalyser.config";
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
		   int sheett = Integer.parseInt(prop.getProperty("worksheetsThreshold"));
		   int fontt = Integer.parseInt(prop.getProperty("fontsThreshold"));
		   int namest = Integer.parseInt(prop.getProperty("definedNamesThreshold"));
		   int cstylest = Integer.parseInt(prop.getProperty("cellStylesThreshold"));
		   int formt = Integer.parseInt(prop.getProperty("formulasThreshold"));
		   int hypert = Integer.parseInt(prop.getProperty("hyperlinksThreshold"));
		   int commt = Integer.parseInt(prop.getProperty("commentsThreshold"));
		   int macrot = Integer.parseInt(prop.getProperty("vbaMacrosThreshold"));
		   int shapet = Integer.parseInt(prop.getProperty("shapesThreshold"));
		   int datet = Integer.parseInt(prop.getProperty("datesThreshold"));
		   int cusedt = Integer.parseInt(prop.getProperty("cellsUsedThreshold"));
		   int extlinkt= Integer.parseInt(prop.getProperty("externalLinksThreshold"));
		   // ... and only then assign them to object properties to ensure use of default values when error
		   worksheetsThreshold = sheett;
		   fontsThreshold = fontt;
		   definedNamesThreshold = namest;
		   cellStylesThreshold = cstylest;
		   formulasThreshold = formt;
		   hyperlinksThreshold = hypert;
		   commentsThreshold = commt;
		   vbaMacrosThreshold = macrot;
		   shapesThreshold = shapet;
		   datesThreshold = datet;
		   cellsUsedThreshold = cusedt;
		   externalLinksThreshold = extlinkt;		   
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
    	options.addOption("c", "config", false, "config file: read complexity assessment threshold values from SpreadsheetComplexityAnalyser.config file");
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
    		printHelpAndExit("Help information for SpreadsheetComplexityAnalyser\n\n"
    				+ "This software extracts values of Excel spreadsheet properties and calculates\n"
    				+ "a tentative spreadsheet complexity assessment based on (default or config\n"
    				+ "file) threshold values.\n\n"
    				+ "The assessment is 'simple/static' or 'complex/dynamic', but feel free to\n"
    				+ "ignore the assessment and use the extracted property values for other purposes.\n\n"
    				+ "This version can extract values for these properties:\n"
    				+ "file: file size, creation date/time, last accessed, last modified\n"
    				+ "workbook: worksheets, fonts, defined names, cell styles, external links and\n"
    				+ "\trevision history"
    				+ "sheet (totaled up): formulas, hyperlinks, comments, shapes, dates, cells used\n"
    				+ "vba: nonzero indicates possible vba macros (tentative)\n\n"
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
            	skipped = processFile(file, true);
            	if (!skipped) {
            		outputResults(file);
            	}
            }
            else if (file.getName().toLowerCase().matches("^(?!\\~\\$)(.*)\\.xl[akms]$")) { // xla xlk xlm xls, no ~$ lock files
            	skipped = processFile(file, false);
            	if (!skipped) {
            		outputResults(file);
            	}
            }
    	}
    	if (xml_out) {
	        System.out.println("<legend>Legend: -1 = not supported (e.g. external links extraction for XLS). 0 or more = number of occurrences. At macros and revision history, nonzero means they are present.</legend>");
    		System.out.println("</spreadsheetComplexityAnalyserResults>");
    	}
    	else if (verbose) {
	        System.out.println("Legend:");
	        System.out.println("\t-1 = not supported (e.g. external links extraction for XLS).");
	        System.out.println("\t0 or more = number of occurrences.");
	        System.out.println("\tAt macros and revision history, nonzero means they are present.");
    	}        
        System.exit(0);
    }
}
