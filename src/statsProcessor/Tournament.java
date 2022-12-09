package statsProcessor;

import java.util.*;
/*
 *
 * THis is a class to hold a tournament 
 */
class Tournament extends Event {
	
	List<Bout> theBouts = new ArrayList<Bout> ();
	
	public void addBout(Bout b) {
		theBouts.add(b);
	}
	public List<Bout> getBouts() {
		return theBouts;
	}
	
	public boolean isATournament() { return true; }
	public boolean isADualMeet() { return false; }
}