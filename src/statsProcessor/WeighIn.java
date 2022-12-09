package statsProcessor;

/*
 *  This is a class to holds an individual WeighIn
 */
class WeighIn {
	
	private float wiWeight;
	private String wiDate;
	private String wiEvent;
	private String entryDateTime;
	
	public float getWIWeight() { return wiWeight; }
	public String getWIDate() { return wiDate; }
	public String getWIEvent() { return wiEvent; }
	public String getEntryDateTime() { return entryDateTime; }
	
	public void setWIWeight(float w) { wiWeight = w; }
	public void setWIDate(String d) { wiDate = d; }
	public void setWIEvent(String e) { wiEvent = e;}
	public void setEntryDateTime(String d) { entryDateTime = d;  }
	
}
