package statsProcessor;

import java.util.*;
import java.text.SimpleDateFormat;


/*
 *  This is a class to holds a history of  individual WeighIn
 */
class WeighInHistory {
	
	private String name;
	private float initialWI=0;
	private String initialWIDate;
	
	private float lastWI=0;
	private String lastWIDate;
	
	private static String INITIAL_WI_TOKEN="Initial Assessment";
	
	
	List<WeighIn> weighIns = new ArrayList<WeighIn> ();
	
	
	public String getName() { return name; }
	public float getInitialWI() { return initialWI; }
	public String getInitialWIDate() { return initialWIDate; }
	public float getLastWI() { return lastWI; }
	public String getLastWIDate() { return lastWIDate; }
	
	public void setName(String n) { name = n; }
	public void setInitialWI(float w) { initialWI = w; }
	public void setInitialWIDate(String d) { initialWIDate = d; }
	public void setLastWI(float lwi) { lastWI = lwi;}
	public void setLastWIDate(String d) { lastWIDate = d; }
	

	public void addWI(WeighIn wi) {
		if ( wi.getWIEvent().contains(INITIAL_WI_TOKEN) ) {
			setInitialWI(wi.getWIWeight());
			setInitialWIDate(wi.getWIDate());
		}
		if ( lastWI == 0 ) {
			setLastWI(wi.getWIWeight());
			setLastWIDate(wi.getWIDate());
		} else {
			try {
				SimpleDateFormat df = new SimpleDateFormat("dd/mm/yyyy");
System.out.println("comparing->" + this.getLastWIDate() + " to " + wi.getWIDate());
				Date newDt = df.parse(wi.getWIDate());
				Date lastWIDt = df.parse(this.getLastWIDate());
				
				if ( lastWIDt.before(newDt) ) {
					setLastWI(wi.getWIWeight());
					setLastWIDate(wi.getWIDate());
				}
			} catch ( Exception e ) {
				System.out.println(" Ran across a bad date in weighin history <" + wi.getWIDate() + ">");
			}
		}
		weighIns.add(wi);
	}
	
    public String findWeighInDate(String event) {
    	if ( event == null ) { return null; }
    	if ( event.length() == 0 ) { return null; }
    	
    	for (int i=0; i < weighIns.size(); i++ ) {
    		WeighIn wi = weighIns.get(i);
    		if ( wi.getWIEvent().equals(event)) {
    			return wi.getWIDate();
    		}
    	}
    	return null;
    }
}

