public class SpreadsheetProperties {
	
	private int iFormulas = 0;
	private int iHyperlinks = 0;
	private int iDefinedNames = 0;
	private int iComments = 0;
	private int iWorkSheets = 0;
	private boolean bVBAMacros = false; // Indicates if a workbook 'seems to have (had)' vba macros
	private int iShapes= 0;
	private int iDates = 0;
	private int iCellsUsed = 0;
	private int iFonts = 0;
	private int iCellStyles = 0;
	private int iColors = 0;

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

	public boolean getHasVBAMacro() {
		return bVBAMacros;
	}

	public void setHasVBAMacros(boolean bVBAMacros) {
		this.bVBAMacros = bVBAMacros;
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

}
