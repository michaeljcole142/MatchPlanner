package statsProcessor;

import java.util.*;
/*
 *
 * THis is a class to hold a team 
 */
class GSTeam {

    private String teamName;
	private List<GSDualMeet> duals = new ArrayList<GSDualMeet> ();
    private List<Tournament> tournaments = new ArrayList<Tournament> ();	
	private List<Tournament> prestigeLastYear = new ArrayList<Tournament> ();
	private List<Tournament> prestige2YearsAgo = new ArrayList<Tournament>();
	private List<Tournament> prestige3YearsAgo = new ArrayList<Tournament>();
 	/*
	 * This roster is key'd on the wrestler's name.
	 */
	private Hashtable<String,GSWrestler> theRoster = new Hashtable<String, GSWrestler> ();

    private Hashtable<String,GSWrestler> lastYearRoster = new Hashtable<String,GSWrestler> ();
	
	private List<GSWrestler>  notWrestlingThisYear;
	
	private Hashtable<String,List<Bout>> allBouts;
	
	
	public String getTeamName() { return teamName; }
	
	public List<Tournament> getPrestigeLastYear() { return prestigeLastYear; }
	public List<Tournament> getPrestige2YearsAgo() { return prestige2YearsAgo; }
	public List<Tournament> getPrestige3YearsAgo() { return prestige3YearsAgo; }
	
	public void addPrestigeLastYearTournament(Tournament t) { prestigeLastYear.add(t); }
	public void addPrestige2YearsAgoTournament(Tournament t) { prestige2YearsAgo.add(t); }
	public void addPrestige3YearsAgoTournament(Tournament t) { prestige3YearsAgo.add(t); }
	
	public Hashtable<String,GSWrestler> getRoster() { return theRoster; }
	public List<GSWrestler> getNotWrestlingThisYear() {
		return notWrestlingThisYear;
	}
	public boolean wrestlerExists(String w) {
		return theRoster.contains(w);
	}
	public boolean wrestlerLastYearExists(String w) {
		return lastYearRoster.contains(w);
	}
	public void addWrestler(GSWrestler w) {
		theRoster.put(w.getName(),w);
	}
	public void addLastYearWrestler(GSWrestler w) {
		lastYearRoster.put(w.getName(),w);
	}
	
	public GSWrestler getWrestler(String n) {
		return theRoster.get(n);
	}
	public GSWrestler getLastYearWrestler(String n) {
		return lastYearRoster.get(n);
	}
	
	public void setTeamName(String name ) {
		teamName = name;
	}
	public int getEventCount() {
		return duals.size() + tournaments.size();
	}
	public void addDualMeet(GSDualMeet d) {
		  duals.add(d);
		  Hashtable<String,Bout> theBouts = d.getBouts();
		  Set<String> keys = theBouts.keySet();
		  for (String key: keys ) {
	
			  Bout b = theBouts.get(key);

			  GSWrestler wrestlerAt = theRoster.get(b.getMainName());

			  if ( wrestlerAt == null ) {
				  GSWrestler newWrestler = new GSWrestler(b.getMainName(),b.getMainTeam());
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
          GSWrestler wrestlerAt = theRoster.get(b.getMainName());
          if ( wrestlerAt == null ) {
            GSWrestler newWrestler = new GSWrestler(b.getMainName(),b.getMainTeam());
			newWrestler.addBout(b);
			theRoster.put(newWrestler.getName(),newWrestler);
		  } else {
            wrestlerAt.addBout(b);
          }	
        }		  
	}	
	public void printVerbose() {
		System.out.println("Team: " + getTeamName() );
		Set<String> keys = theRoster.keySet();
		for (String key: keys ) {
			GSWrestler wrestlerAt = theRoster.get(key);
			wrestlerAt.printVerbose();
		}
	}
	public void buildAllBoutsLookup() {
		allBouts = new Hashtable<String,List<Bout>> ();
		Set<String> keys = theRoster.keySet();
		for (String key: keys ) {
			GSWrestler wrestlerAt = theRoster.get(key);
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

			GSWrestler lastYearWrestler = lastYearRoster.get(key);
			GSWrestler atWrestler = theRoster.get(key);
			if ( lastYearWrestler != null ) {
				atWrestler = theRoster.get(key);
				atWrestler.setRosteredLastYearOn();
				atWrestler.setTrackLastYearRecord(lastYearWrestler.getTrackRecord());
				atWrestler.setTrackLastYearWtClass(lastYearWrestler.getTrackWtClass());
			} else {
				atWrestler.setRosteredLastYearOff();	
			}
		}
		notWrestlingThisYear = new ArrayList<GSWrestler> ();
		keys = lastYearRoster.keySet();
		for ( String key: keys ) {
			GSWrestler w = theRoster.get(key);
			
			if ( w == null ) {
				GSWrestler nRostered = lastYearRoster.get(key);
				if ( nRostered != null ) {
			
					if ( nRostered.getGrade() != WrestlingLanguage.Grade.G7 ) {
						System.out.println("At NR wrestler ");
						nRostered.printVerbose();
						notWrestlingThisYear.add(nRostered);
					}
				}
			}
		}
    }
	
}