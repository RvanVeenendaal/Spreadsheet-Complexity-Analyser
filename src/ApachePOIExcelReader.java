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
import java.util.Collection;
import java.util.Iterator;
import java.util.Map;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.Option;
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
    private static SpreadsheetProperties sp;

    public static void processFile(File file, boolean isXSSF) {    
    	try {
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
	        	sp.setExternalLinks(-1); // not available for hssf
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
	    } catch (FileNotFoundException e) {
	        e.printStackTrace();
	    } catch (IOException e) {
	        e.printStackTrace();
	    } catch (Exception e) {
	    	e.printStackTrace();
	    }
    	try {
    		FileInputStream excelFile = new FileInputStream(file);
	        VBAMacroReader reader = null;
	        Map<String, String> macros = null;
            reader = new VBAMacroReader(excelFile);
            macros = reader.readMacros();
            /* macros contains modules that might be empty
             * macros also contains Excel objects like Sheet1 and ThisWorkbook
             * therefore, we cannot be sure that there actually is a vba macro
            */
            if (macros.size() > 0) { // due to Sheet1 and ThisWorkbook, perhaps > 2 ?
                sp.setHasVBAMacros(true);
            }            
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

    
// Todo: read threshold values for result calculation from config.ini file (or as CLI parameters?)    
// and use those values for calculating and outputting result simple/static or complex/dynamic
// but only when user asks for result calculation via CLI parameter
    public static void outputResults() {   
		String result = null;
	    if (sp.getWorkSheets() > 1 ||
	    		sp.getFonts() > 1 ||
	    		sp.getDefinedNames() > 0 ||
	    		sp.getCellStyles() > 1 ||
	    		sp.getFormulas() > 0 ||
	    		sp.getHyperlinks() > 0 ||
	    		sp.getComments() > 0 ||
	    		sp.getHasVBAMacro() == true ||
	    		sp.getShapes() > 0 ||
	    		sp.getDates() > 0 ||
	    		sp.getCellsUsed() > 1000) {
	    	result = "complex/dynamic";
	    }
	    else {
	    	result = "simple/static";
	    }    	
    	if (xml_out) {
	        System.out.println("<spreadsheetComplexityReport'>");
//	        System.out.println("\t<result>" + result + "</result>");
	        System.out.println("\t<worksheets>" + sp.getWorkSheets() + "</worksheets>");
	        System.out.println("\t<fonts>" + sp.getFonts() + "</>");
	        System.out.println("\t<definedNames>" + sp.getDefinedNames() + "</definedNames>");
	        System.out.println("\t<cellStyles>" + sp.getCellStyles() + "</cellStyles>");
	        System.out.println("\t<formulas>" + sp.getFormulas() + "</formulas>");
	        System.out.println("\t<hyperlinks>" + sp.getHyperlinks() + "</hyperlinks>");
	        System.out.println("\t<comments>" + sp.getComments() + "</comments>");
	        System.out.println("\t<vbaMacros>" + sp.getHasVBAMacro() + "</vbaMacros>");
	        System.out.println("\t<shapes>" + sp.getShapes() + "</shapes>");
	        System.out.println("\t<dates>" + sp.getDates() + "</dates>");
	        System.out.println("\t<usedCells>" + sp.getCellsUsed() + "</usedCells>");
	        System.out.println("\t<externalLinks>" + sp.getExternalLinks() + "</externalLinks>");
	        System.out.println("</spreadsheetComplexityReport>");
    	}
    	else if (verbose) {
//    	    System.out.println("Result: " + result);
    		System.out.println("Spreadsheet complexity analyser results:");
	        System.out.println("Number of");
	        System.out.println("\tworksheets:\t\t" + sp.getWorkSheets());
	        System.out.println("\tfonts:\t\t\t" + sp.getFonts());
	        System.out.println("\tdefined names:\t\t" + sp.getDefinedNames());
	        System.out.println("\tcell styles:\t\t" + sp.getCellStyles()); 
	        System.out.println("\tformulas:\t\t" + sp.getFormulas());
	        System.out.println("\thyperlinks:\t\t" + sp.getHyperlinks());
	        System.out.println("\tcomments:\t\t" + sp.getComments());
	        System.out.println("\tvba macros:\t\t" + sp.getHasVBAMacro());
	        System.out.println("\tshapes:\t\t\t" + sp.getShapes());
	        System.out.println("\tdates:\t\t\t" + sp.getDates());
	        System.out.println("\tcells used:\t\t" + sp.getCellsUsed());
	        System.out.println("\texternal links:\t\t" + sp.getExternalLinks());
	        System.out.println("Legend:");
	        System.out.println("\t0 or more = number of occurrences,");
	        System.out.println("\tfalse = does not occur (for VBA macros),");
	        System.out.println("\ttrue = probably occurs (for VBA macros).");
		}
    	else {
    	    System.out.println("Spreadsheet complexity analyser result summary: probably " + result);
    	}
    }
     
/*
 * Spreadsheet Complexity Analyser
 * java -jar spreadsheet-complexity-analyser.jar folder file(s) >
 * 	outputs filename and conclusion ("simple/static" or "complex/dynamic")
 * java -jar spreadsheet-complexity-analyser.jar -r folder file(s)
 *  same, but recurses subfolders
 * java -jar spreadsheet-complexity-analyser.jar -v folder file(s)
 *  same as first, but also outputs number of occurrences of properties
 * java -jar spreadsheet-complexity-analyser.jar -v -r folder file(s)
 *  same as first, but recurses subfolders and outputs number of occurrences
 */
    
    public static void main(String[] args) throws IOException {
    	CommandLineParser parser = new DefaultParser();
    	HelpFormatter formatter = new HelpFormatter();
    	CommandLine cmd = null;
    	Options options = new Options();
    	options.addOption("v", "verbose", false, "verbose output: show number of occurrences of properties in text form" );
    	options.addOption("x", "xml_out", false, "xml output: show number of occurrences of properties in xml form (suppresses verbose output)");
    	options.addOption("r", "recursive", false, "recurse into subdirectories" );
    	options.addOption( Option.builder( "dir" )
    			.required(true)
    			.hasArg()
    			.desc("directory to process" )
    			.build());
    	options.addOption( Option.builder( "file" )
    			.required(true)
    			.hasArg()
    			.desc("filename (wildcards allowed)" )
    			.build());
    	try {
    	    // parse the command line arguments
    	    cmd = parser.parse( options, args );
    	}
    	catch( ParseException exp ) {
    	    System.out.println( "Unexpected exception:" + exp.getMessage() );
    	    formatter.printHelp("spreadsheet-complexity-analyser", options);
    	    System.exit(0);
    	}

    	verbose = cmd.hasOption("verbose");
    	xml_out = cmd.hasOption("xml_out");
    	recursive = cmd.hasOption("recursive");
    	String dirName = cmd.getOptionValue("dir");
    	String fileName = cmd.getOptionValue("file");
    	
    	File dir = new File(dirName);
    	WildcardFileFilter fileFilter = new WildcardFileFilter(fileName);
    	Collection<File> files = null;
    	if (recursive) {
    		files = FileUtils.listFiles(dir, fileFilter, DirectoryFileFilter.DIRECTORY);        	
    	}
    	else {
    		files = FileUtils.listFiles(dir, fileFilter, null);
    	}
    	Iterator<File> fileIterator = files.iterator();
        while (fileIterator.hasNext()) {
        	sp = new SpreadsheetProperties();
            File file = fileIterator.next();
        	System.out.println("Processing " + file);
            if (file.getName().toLowerCase().matches("(.*)\\.xl[st][xm]$")) { //xlsx xlsm xltx xltm
            	processFile(file, true);
                outputResults();
            }
            else if (file.getName().toLowerCase().matches("(.*)\\.xl[akms]$")) {
            	processFile(file, false);
                outputResults();
            }
            else {
            	System.out.println("Error: file " + file.getName() + " is not an Excel file.");
            }
    	}
        System.exit(0);
    }
}
