public class SpreadsheetProperties {
	
	private int iFormulas = 0; // formulas present in spreadsheet
	private int iHyperlinks = 0; // hyperlinks present in spreadsheet
	private int iDefinedNames = 0; // defined names present in spreadsheet
	private int iComments = 0; // comments present in spreadsheet
	private int iWorkSheets = 0; // worksheets present in spreadsheet
	private int iHasVBAMacros = 0; // spreadsheet has (had) vba macros: 0 = false, 1 or greater = true
	private int iShapes= 0; // shapes present in spreadsheet
	private int iDates = 0; // dates present in spreadsheet
	private int iCellsUsed = 0; // cells used in spreadsheet
	private int iFonts = 0; // fonts used in spreadsheet
	private int iCellStyles = 0; // cell styles used in spreadsheet
	private int iColors = 0; // colours used in spreadsheet
	private int iExternalLinks = 0; // external links present in formulas (!) in spreadsheet
	private double iFileSizeKb = 0; // size of file (in kilobyte)
	private String sLastModified = ""; // last modified time (if available) 
	private String sLastAccessed = ""; // last accessed time (if available)
	private String sCreation = ""; // creation time (if available)
	private int iHasRevisionHistory = 0; // 1 = spreadsheet has revision history on, 0 = off

	public int getFormulas() {
		return iFormulas;
	}

	public void setFormulas(int iFormulas) {
		this.iFormulas = iFormulas;
	}

	public int getHyperlinks() {
		return iHyperlinks;
	}

	public void setHyperlinks(int iHyperlinks) {
		this.iHyperlinks = iHyperlinks;
	}

	public int getDefinedNames() {
		return iDefinedNames;
	}

	public void setDefinedNames(int iDefinedNames) {
		this.iDefinedNames = iDefinedNames;
	}

	public int getComments() {
		return iComments;
	}

	public void setComments(int iComments) {
		this.iComments = iComments;
	}

	public int getWorkSheets() {
		return iWorkSheets;
	}

	public void setWorkSheets(int iWorkSheets) {
		this.iWorkSheets = iWorkSheets;
	}

	public int getHasVBAMacros() {
		return iHasVBAMacros;
	}

	public void setHasVBAMacros(int iHasVBAMacros) {
		this.iHasVBAMacros = iHasVBAMacros;
	}

	public int getShapes() {
		return iShapes;
	}

	public void setShapes(int iShapes) {
		this.iShapes = iShapes;
	}

	public int getDates() {
		return iDates;
	}

	public void setDates(int iDates) {
		this.iDates = iDates;
	}

	public int getCellsUsed() {
		return iCellsUsed;
	}

	public void setCellsUsed(int iCellsUsed) {
		this.iCellsUsed = iCellsUsed;
	}

	public int getFonts() {
		return iFonts;
	}

	public void setFonts(int iFonts) {
		this.iFonts = iFonts;
	}

	public int getColors() {
		return iColors;
	}

	public void setColors(int iColors) {
		this.iColors = iColors;
	}

	public int getCellStyles() {
		return iCellStyles;
	}

	public void setCellStyles(int iCellStyles) {
		this.iCellStyles = iCellStyles;
	}

	public int getExternalLinks() {
		return iExternalLinks;
	}

	public void setExternalLinks(int iExternalLinks) {
		this.iExternalLinks = iExternalLinks;
	}

	public double getFileSizeKb() {
		return iFileSizeKb;
	}

	public void setFileSizeKb(double iFileSizeKb) {
		this.iFileSizeKb = iFileSizeKb;
	}

	public String getLastModified() {
		return sLastModified;
	}

	public void setLastModified(String sLastModified) {
		this.sLastModified = sLastModified;
	}

	public String getLastAccessed() {
		return sLastAccessed;
	}

	public void setLastAccessed(String sLastAccessed) {
		this.sLastAccessed = sLastAccessed;
	}

	public String getCreation() {
		return sCreation;
	}

	public void setCreation(String sCreation) {
		this.sCreation = sCreation;
	}

	public int getHasRevisionHistory() {
		return iHasRevisionHistory;
	}

	public void setHasRevisionHistory(int iHasRevisionHistory) {
		this.iHasRevisionHistory = iHasRevisionHistory;
	}
}
