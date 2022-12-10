public class FileProperties {
	private double dFileSizeKb = 0; // size of file (in kilobyte)
	private String sLastModified = ""; // last modified time (if available) 
	private String sLastAccessed = ""; // last accessed time (if available)
	private String sCreation = ""; // creation time (if available)

	public double getdFileSizeKb() {
		return dFileSizeKb;
	}
	public void setdFileSizeKb(double dFileSizeKb) {
		this.dFileSizeKb = dFileSizeKb;
	}
	public String getsLastModified() {
		return sLastModified;
	}
	public void setsLastModified(String sLastModified) {
		this.sLastModified = sLastModified;
	}
	public String getsLastAccessed() {
		return sLastAccessed;
	}
	public void setsLastAccessed(String sLastAccessed) {
		this.sLastAccessed = sLastAccessed;
	}
	public String getsCreation() {
		return sCreation;
	}
	public void setsCreation(String sCreation) {
		this.sCreation = sCreation;
	}
}
