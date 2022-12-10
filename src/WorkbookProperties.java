public class WorkbookProperties {
	
	private int iDefinedNames = 0; // defined names present in spreadsheet
	private int iWorkSheets = 0; // worksheets present in spreadsheet
	private int iFonts = 0; // fonts used in spreadsheet
	private int iCellStyles = 0; // cell styles used in spreadsheet
	private int iExternalLinks = 0; // external links present in formulas (!) in spreadsheet
	private int iVBAMacros = 0; // spreadsheet has (had) vba macros: 0 = false, 1 or greater = true
	private int iHasRevisionHistory = 0; // 1 = spreadsheet has revision history on, 0 = off, -1 is n/a
	
	public int getiHasRevisionHistory() {
		return iHasRevisionHistory;
	}
	public void setiHasRevisionHistory(int iHasRevisionHistory) {
		this.iHasRevisionHistory = iHasRevisionHistory;
	}
	public int getiVBAMacros() {
		return iVBAMacros;
	}
	public void setiVBAMacros(int iVBAMacros) {
		this.iVBAMacros = iVBAMacros;
	}
	public int getiDefinedNames() {
		return iDefinedNames;
	}
	public void setiDefinedNames(int iDefinedNames) {
		this.iDefinedNames = iDefinedNames;
	}
	public int getiWorkSheets() {
		return iWorkSheets;
	}
	public void setiWorkSheets(int iWorkSheets) {
		this.iWorkSheets = iWorkSheets;
	}
	public int getiFonts() {
		return iFonts;
	}
	public void setiFonts(int iFonts) {
		this.iFonts = iFonts;
	}
	public int getiCellStyles() {
		return iCellStyles;
	}
	public void setiCellStyles(int iCellStyles) {
		this.iCellStyles = iCellStyles;
	}
	public int getiExternalLinks() {
		return iExternalLinks;
	}
	public void setiExternalLinks(int iExternalLinks) {
		this.iExternalLinks = iExternalLinks;
	}
	
}
