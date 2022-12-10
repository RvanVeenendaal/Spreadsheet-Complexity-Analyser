import java.util.ArrayList;

public class SpreadsheetProperties {
	private FileProperties fileProperties = new FileProperties(); 					// A spreadsheet has one set of file properties
	private WorkbookProperties workbookProperties = new WorkbookProperties();			// A spreadsheet has one set of workbook properties
	private ArrayList<WorksheetProperties> worksheetPropertiesList = new ArrayList<>();		// A spreadsheet (workbook) can have multiple work sheets
	
	public ArrayList<WorksheetProperties> getWorksheetPropertiesList() {
		return worksheetPropertiesList;
	}
	public void setWorksheetPropertiesList(ArrayList<WorksheetProperties> worksheetPropertiesList) {
		this.worksheetPropertiesList = worksheetPropertiesList;
	}
	public FileProperties getFileProperties() {
		return fileProperties;
	}
	public void setFileProperties(FileProperties fileProperties) {
		this.fileProperties = fileProperties;
	}
	public WorkbookProperties getWorkbookProperties() {
		return workbookProperties;
	}
	public void setWorkbookProperties(WorkbookProperties workbookProperties) {
		this.workbookProperties = workbookProperties;
	}
}