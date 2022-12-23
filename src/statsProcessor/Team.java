package statsProcessor;

import java.util.*;
/*
 *
 * THis is a class to hold a team 
 */
class Team {

    private String teamName;
	private List<DualMeet> duals = new ArrayList<DualMeet> ();
	private List<DualMeet> lastYearDuals = new ArrayList<DualMeet> ();

	private List<Tournament> tournaments = new ArrayList<Tournament> ();
    private List<Tournament> lastYearTournaments = new ArrayList<Tournament> ();	
	private List<Tournament> prestigeLastYear = new ArrayList<Tournament> ();
	private List<Tournament> prestige2YearsAgo = new ArrayList<Tournament>();
	private List<Tournament> prestige3YearsAgo = new ArrayList<Tournament>();
	private List<Bout> lastYearBouts = new ArrayList<Bout> ();

	/*
	 * This roster is key'd on the wrestler's name.
	 */
	private Hashtable<String,Wrestler> theRoster = new Hashtable<String, Wrestler> ();

    private Hashtable<String,Wrestler> lastYearRoster = new Hashtable<String,Wrestler> ();
	
	private List<Wrestler>  notWrestlingThisYear;
	
	private Hashtable<String,List<Bout>> allBouts;
	
	
	public String getTeamName() { return teamName; }
	
	public List<Tournament> getPrestigeLastYear() { return prestigeLastYear; }
	public List<Tournament> getPrestige2YearsAgo() { return prestige2YearsAgo; }
	public List<Tournament> getPrestige3YearsAgo() { return prestige3YearsAgo; }
	
	public void addPrestigeLastYearTournament(Tournament t) { prestigeLastYear.add(t); }
	public void addPrestige2YearsAgoTournament(Tournament t) { prestige2YearsAgo.add(t); }
	public void addPrestige3YearsAgoTournament(Tournament t) { prestige3YearsAgo.add(t); }
	
	public Hashtable<String,Wrestler> getRoster() { return theRoster; }
	public List<Wrestler> getNotWrestlingThisYear() {
		return notWrestlingThisYear;
	}
	public boolean wrestlerExists(String w) {
		return theRoster.contains(w);
	}
	public boolean wrestlerLastYearExists(String w) {
		return lastYearRoster.contains(w);
	}
	public void addWrestler(Wrestler w) {
		theRoster.put(w.getName(),w);
	}
	public void addLastYearWrestler(Wrestler w) {
		lastYearRoster.put(w.getName(),w);
	}
	
	public Wrestler getWrestler(String n) {
		return theRoster.get(n);
	}
	public Wrestler getLastYearWrestler(String n) {
		return lastYearRoster.get(n);
	}
	
	public void setTeamName(String name ) {
		teamName = name;
	}
	public int getEventCount() {
		return duals.size() + tournaments.size();
	}
	public void addDualMeet(DualMeet d) {
		  duals.add(d);
		  Hashtable<String,Bout> theBouts = d.getBouts();
		  Set<String> keys = theBouts.keySet();
		  for (String key: keys ) {
	
			  Bout b = theBouts.get(key);

			  Wrestler wrestlerAt = theRoster.get(b.getMainName());

			  if ( wrestlerAt == null ) {
				  Wrestler newWrestler = new Wrestler(b.getMainName(),b.getMainTeam());
				  newWrestler.addBout(b);
				  theRoster.put(newWrestler.getName(),newWrestler);
		      } else {
				  wrestlerAt.addBout(b);
			  }
		  }
    }
	public void addTournament(Tournament t) {
		tournaments.add(t);
        List<Bout> theBouts = t.getBouts();
        for ( int i=0; i < theBouts.size() ; i++ ) {
        	Bout b = theBouts.get(i);
        	Wrestler wrestlerAt = theRoster.get(b.getMainName());
        	if ( wrestlerAt == null ) {
        		Wrestler newWrestler = new Wrestler(b.getMainName(),b.getMainTeam());
        		newWrestler.addBout(b);
        		theRoster.put(newWrestler.getName(),newWrestler);
        	} else {
        		wrestlerAt.addBout(b);
        	}	
        }		  
	}
	public void addLYTournament(Tournament t) {
		this.lastYearTournaments.add(t);
        List<Bout> theBouts = t.getBouts();
        for ( int i=0; i < theBouts.size() ; i++ ) {
        	Bout b = theBouts.get(i);
        	Wrestler wrestlerAt = theRoster.get(b.getMainName());
        	if ( wrestlerAt != null ) {
        		wrestlerAt.addLYBout(b);
        	}
        	this.lastYearBouts.add(b);
        }		 
	}
	public void addLYDualMeet(DualMeet d) {
		  this.lastYearDuals.add(d);
		  Hashtable<String,Bout> theBouts = d.getBouts();
		  Set<String> keys = theBouts.keySet();
		  for (String key: keys ) {
	
			  Bout b = theBouts.get(key);

			  Wrestler wrestlerAt = theRoster.get(b.getMainName());

			  if ( wrestlerAt != null ) {
				  wrestlerAt.addLYBout(b);
			  }
			  this.lastYearBouts.add(b);
		  }
	}
	public String processLastYearBouts(String opponentTeam, String opponentName) {
		String ret="";
		for ( int i=0; i < this.lastYearBouts.size(); i++ ) {
			Bout b = this.lastYearBouts.get(i);
			if ( b.getOpponentTeam().equals(opponentTeam) && 
					b.getOpponentName().equals(opponentName)) {
				System.out.println("LastYearBout->" + b.toString());
				if (ret.length() > 0 ) { ret +="\n";}
				ret += "LAST YEAR: " + b.toString();
			}
		}
		return ret;
	}
	public void printVerbose() {
		System.out.println("Team: " + getTeamName() );
		Set<String> keys = theRoster.keySet();
		for (String key: keys ) {
			Wrestler wrestlerAt = theRoster.get(key);
			wrestlerAt.printVerbose();
		}
	}
	public void buildAllBoutsLookup() {
		allBouts = new Hashtable<String,List<Bout>> ();
		Set<String> keys = theRoster.keySet();
		for (String key: keys ) {
			Wrestler wrestlerAt = theRoster.get(key);
			List<Bout> bouts = wrestlerAt.getBouts();
			for ( int i=0; i < bouts.size(); i++ ) {
				Bout b = bouts.get(i);
				if ( b.getMatchResultType() != WrestlingLanguage.MatchResultType.FFT ) {
					
					String oppKey = b.getOpponentTeam() + ":" + b.getOpponentName();
					List<Bout> bList = allBouts.get(oppKey);
					if ( bList == null ) {
						bList = new ArrayList<Bout> ();
						bList.add(b);
						allBouts.put(oppKey,bList);
					} else {
						bList.add(b);
						allBouts.replace(oppKey,bList);
					}
				}
			}
		}
	}
	
	/* The key here is opponentTeam + ":" + opponentName */
	
	public List<Bout> lookupCommonMatch (String key) {
		return allBouts.get(key);
	}
	
	public void syncWithLastYearRoster() {
			 
		Set<String> keys = theRoster.keySet();
		for (String key: keys ) {

			Wrestler lastYearWrestler = lastYearRoster.get(key);
			Wrestler atWrestler = theRoster.get(key);
			if ( lastYearWrestler != null ) {
				atWrestler = theRoster.get(key);
				atWrestler.setRosteredLastYearOn();
				atWrestler.setTrackLastYearRecord(lastYearWrestler.getTrackRecord());
				atWrestler.setTrackLastYearWtClass(lastYearWrestler.getTrackWtClass());
			} else {
				atWrestler.setRosteredLastYearOff();	
			}
		}
		notWrestlingThisYear = new ArrayList<Wrestler> ();
		keys = lastYearRoster.keySet();
		for ( String key: keys ) {
			Wrestler w = theRoster.get(key);
			
			if ( w == null ) {
				Wrestler nRostered = lastYearRoster.get(key);
				if ( nRostered != null ) {
			
					if ( nRostered.getGrade() != WrestlingLanguage.Grade.SR ) {
						System.out.println("At NR wrestler ");
						nRostered.printVerbose();
						notWrestlingThisYear.add(nRostered);
					}
				}
			}
		}
    }
	
}