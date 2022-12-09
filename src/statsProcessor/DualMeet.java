package statsProcessor;

import java.util.*;



/*
 *  This is a class to hold dual meets.
 *  Notice that it extends Event.  An event is 
 *  a class that holds the things that are common to both a
 *  Dual Meet and a Tournament.
 */
class DualMeet extends Event {
	
	
	private List<String> weights = new ArrayList<String> ();
	
	private String opponent;
	/*
	 * This Hashtable is key'd on Weight.
	 */
	private Hashtable<String,Bout>  theDual = new Hashtable<String,Bout>();
	
	public List<String> getWeights() { return weights; }

	public boolean isADualMeet() { return true; }
	public boolean isATournament() { return false; }
	
	private String[] HSWeightList= {"106","113","120","126","132","138","144","150","157","165","175","190","215","285"};
	
	DualMeet() {
		//defaults to HS weights.  
		this.initialize(HSWeightList);
	}
	protected void resetWeights() { 
		this.weights = new ArrayList<String> ();
	}
	public void initialize(String[] wts) {
		for ( int i=0; i < wts.length; i++) {
			this.weights.add(wts[i]);
		}
	}
	
	public int getDualSize() { 
		return this.weights.size();
	}
	public String getOpponent() { return opponent; }
	
	public void setOpponent( String opp ) {
		opponent = opp;
	}
   
    public Hashtable<String,Bout> getBouts() {
		return theDual;
	}
	public String toString() {
		return getMainTeam() + ":" + getOpponent() + ":" + getEventDate();
	}
	public void addBout(Bout b) {
		theDual.put(b.getWeight(),b);
	}
	public Bout getBout(String w) {
		return theDual.get(w);
	}
	public boolean doesBoutExist(String w) {
	    return theDual.contains(w);
	}
	public boolean isFullDual() {
		if ( theDual.size() != weights.size() ) {
			return false;
		}
		
		
		Set<String> keys = theDual.keySet();
		
		for (String key: keys ) {
			weights.remove(key);
		}
		if ( weights.size() > 0 ) {
			return false;
		}
		return true;
	}
}
