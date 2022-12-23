package statsProcessor;

import java.util.*;


/*
 *
 *  This is a class to holds an wrestler and all data pertaining to him/her.
 */
class Wrestler {
	private String name;
	private String teamName;
	private WrestlingLanguage.Grade grade;
	private WrestlingLanguage.Gender gender;
	private String trackRecord;
	private String trackWtClass;
	private boolean rosteredLastYear;
	private String trackLastYearWtClass;
	private String trackLastYearRecord="";
	private WrestlingLanguage.Prestige prestigeLastYear;
	private WrestlingLanguage.Prestige prestige2YearsAgo;
	private WrestlingLanguage.Prestige prestige3YearsAgo;
	
	private Hashtable<String,Integer> matchesAtWeight=new Hashtable<String,Integer>();
	
	private int wins;
	private int losses;
	private boolean lossByInjury=false;
	private String lossByInjuryString;
	private int winByFFT;

	private int winByFall;
	private int winByDec;
	private int winByMD;
	private int winByTech;
	
	private int lossByFall;
	private int lossByDec;
	private int lossByMD;
	private int lossByTech;
	
	private WeighInHistory wiHistory;
	
	/* These 2 attributes are used to force where we put the wrestler in the reporting
	 * order.
	 */
	private String forceWeight=null;  /* for reporting force the use of this weight class. */
	private Integer forceRank=null;  /* for reporting this is the rank order withing the weight class. */

	public void setLossByInjuryString(String s) {
		lossByInjuryString = s;
	}
	public String getLossByInjuryString() { return lossByInjuryString; }
	
	public String getLowestWeightWrestled() {
		int lowInt=9999;
		
		Set<String> keys = matchesAtWeight.keySet();
		for (String key: keys ) {
			int i =  Integer.parseInt(key);
		    if (i < lowInt ) {
		    	lowInt = i;
		    }
		}
		if ( lowInt != 9999 ) {
			return Integer.toString(lowInt);
		}
		if ( this.trackWtClass != null && this.trackWtClass.length() > 0 ) {
			return this.trackWtClass;
		}
		return "9999";
	}
	public String getPrintWeight() {
		if ( forceWeight != null ) {
			if ( forceWeight.length() > 0 ) {
				return forceWeight;
			}
		}
		return getLowestWeightWrestled();
	}
	public int getPrintRank() {
		if (forceRank == null ) {
			return getWins() + getLosses();
		} else {
			return 100000 - forceRank.intValue();
		}
	}
	public String getForceWeight() { return forceWeight; }
	public Integer getForceRank() { return forceRank; }
	
	public void setForceWeight( String w ) { forceWeight = w; }
	public void setForceRank ( Integer i ) { forceRank = i; }
	
	List<Bout> bouts = new ArrayList<Bout> ();
	List<Bout> lastYearBouts = new ArrayList<Bout> ();

	public List<Bout> getBouts() { return bouts; }
	
	public boolean getLossByInjury() { return lossByInjury; }

	public List<Bout> getLastYearBouts() { return this.lastYearBouts;}
	
	public WeighInHistory getWeighInHistory() { return wiHistory; }
	
	public void setWeighInHistory(WeighInHistory wih ) { wiHistory = wih; }

	public void setPrestigeLastYear(WrestlingLanguage.Prestige p) { prestigeLastYear = p; }
	public void setPrestige2YearsAgo(WrestlingLanguage.Prestige p) { prestige2YearsAgo= p; }
	public void setPrestige3YearsAgo(WrestlingLanguage.Prestige p) { prestige3YearsAgo = p; }

	public WrestlingLanguage.Prestige getPrestigeLastYear() { return prestigeLastYear; }
	public WrestlingLanguage.Prestige getPrestige2YearsAgo() { return prestige2YearsAgo; }
	public WrestlingLanguage.Prestige getPrestige3YearsAgo() { return prestige3YearsAgo; }

    public void printVerbose () {
		System.out.println( getName() + 
			" Grade<" + getGrade() + ">" + 
			" Gender<" + getGender() + "> Wt<" + 
			getTrackWtClass() + ">" + " (" + getWins() + "W-" + 
			getLosses() + "L) Last Year=<" + rosteredLastYear + "> Last Year Record <" + trackLastYearRecord + ">" );
    }
  
	public String getRecordString() { 
		return String.valueOf(wins) + "-" + String.valueOf(losses); 
	}
	public String getRecordBreakdown() {
		String breakdownString = "";
		boolean addedWins=false;
		if ( getWins() > 0 ) {
			breakdownString += getWins() + "W[";
			if ( getWinByFall() > 0 ) {
				breakdownString += getWinByFall() + " Pin,";
				addedWins=true;
			}
			if ( getWinByTech() > 0 ) {
				breakdownString += getWinByTech() + " TF,";
				addedWins=true;
			}
			if ( getWinByMD() > 0 ) {
				breakdownString += getWinByMD() + " MD,";
				addedWins=true;
			}
			if ( getWinByDec() > 0 ) {
				breakdownString += getWinByDec() + " Dec,";
				addedWins=true;
			}
			if ( getWinByFFT() > 0 ) {
				breakdownString += getWinByFFT() + " FFT,";
				addedWins=true;
			}
			if ( addedWins ) { 
				breakdownString = breakdownString.substring(0,breakdownString.length()-1);
			}
			breakdownString += "]";
		} else {
			breakdownString += "0 W";
		}
		boolean addedLosses=false;
		if ( getLosses() > 0 ) {
			breakdownString += "-" + getLosses() + "L[";
			if ( getLossByFall() > 0 ) {
				breakdownString += getLossByFall() + " Pin,";
				addedLosses=true;
			}
			if ( getLossByTech() > 0 ) {
				breakdownString += getLossByTech() + " TF,";
				addedLosses=true;
			}
			if ( getLossByMD() > 0 ) {
				breakdownString += getLossByMD() + " MD,";
				addedLosses=true;
			}
			if ( getLossByDec() > 0 ) {
				breakdownString += getLossByDec() + " Dec,";
				addedLosses=true;
			}
			
			if ( addedLosses ) { 
				breakdownString = breakdownString.substring(0,breakdownString.length()-1);
			}
			breakdownString += "]";
		} else {
			breakdownString += "-0L";
		}
		if ( getWins() ==0 && getLosses() == 0 ) {
			return "";
		} else {
			return breakdownString;
		}
	}
	public String getName() { return name; }
	public String getTeamName() { return teamName; }
	public WrestlingLanguage.Grade getGrade() { return grade; }
	public WrestlingLanguage.Gender getGender() { return gender; }
	public String getTrackRecord() { return trackRecord; }
	public String getTrackWtClass() { return trackWtClass; }
	
	public int getWins() { return wins; }
	public int getLosses() { return losses; }
	
	public int getWinByFFT() { return winByFFT; }
	public int getWinByFall() { return winByFall; }
	public int getWinByTech() { return winByTech; }
	public int getWinByMD() { return winByMD; }
	public int getWinByDec() { return winByDec; }
	
	public int getLossByFall() { return lossByFall; }
	public int getLossByTech() { return lossByTech; }
	public int getLossByMD() { return lossByMD; }
	public int getLossByDec() { return lossByDec; }
	
    public Wrestler(String n, String team){
		setName(n);
		setTeamName(team);
		return;
	}
   
	public boolean getRosteredLastYear() { return rosteredLastYear; }
	
	public void setRosteredLastYearOn() { rosteredLastYear = true; }
	public void setRosteredLastYearOff() { rosteredLastYear = false; }
	
	public String getTrackLastYearWtClass() { return trackLastYearWtClass; }
	public String getTrackLastYearRecord() { return trackLastYearRecord; }
	
	public void setTrackLastYearWtClass(String w) { 
		trackLastYearWtClass = w; 
	}
	public void setTrackLastYearRecord(String r) {
		trackLastYearRecord = r;
	}
	
	
	public void addBout(Bout b) {
		String cmd=b.getMatchDate();
		if (cmd == null ) {
			Event e = b.getEvent();
			if ( e != null ) {
				String sss = e.getEventTitle();
				if ( wiHistory != null ) {
					String s = wiHistory.findWeighInDate(sss);
					if ( s != null ) { 
						b.setMatchDate(s);
					}
				}
			}
		}	
		
		bouts.add(b);
		addToSummary(b);
		addToMatchesAtWeight(b);
	}
	public void addLYBout(Bout b) {
		lastYearBouts.add(b);
	}
	
	private void addToMatchesAtWeight(Bout b) {
		String w=b.getWeight();
		Integer ct = matchesAtWeight.get(w);
		if ( ct != null ) {
			ct++;
			matchesAtWeight.replace(w,ct);
		} else {
			matchesAtWeight.put(w,1);
		}
		return;
	}
	public String getMatchesAtWeightString() {

		String byWt="";
		Set<String> keys = matchesAtWeight.keySet();
		for (String key: keys ) {
	
			int ct = matchesAtWeight.get(key);
			if ( byWt.length() > 0 ) { byWt=byWt + ";" ; }
			byWt=byWt + ct + "@" + key;
		}
		return byWt;
	}
	
	public void setName( String n ) {
		name = n.strip();
	}
	public void setTeamName( String team ) {
		teamName = team;
	}
	public void setGradeFr() { grade = WrestlingLanguage.Grade.FR;	}
	public void setGradeSo() { grade = WrestlingLanguage.Grade.SO;	}
	public void setGradeJr() { grade = WrestlingLanguage.Grade.JR;	}
	public void setGradeSr() { grade = WrestlingLanguage.Grade.SR;	}
	public void setGrade(WrestlingLanguage.Grade g ) { grade=g; }
	
	public void setGenderM() { gender = WrestlingLanguage.Gender.M;	}
	public void setGenderF() { gender = WrestlingLanguage.Gender.F;	}
	public void setGender(WrestlingLanguage.Gender g ) { gender = g; }
	
	public void setTrackRecord(String rec) {
		trackRecord = rec;
	}
	public void setTrackWtClass(String w ) {
		trackWtClass = w;
	}
	public String toString() {
		return "Wrestler:" + teamName + ":" + name  ;
	}
	protected void addToSummary(Bout b) {
	   if ( b.isAWin() ) {
          addWin();	
          switch ( b.getMatchResultType() ) {
			  case FFT: 
			    addWinByFFT();
				break;
			  case FALL:
			    addWinByFall();
				break;
			  case TECH:
			    addWinByTech();
				break;
			  case MD:
			    addWinByMD();
			    break;
			  case DEC:
			    addWinByDec();
				break;
			  case INJ:
			    break;
			  case DQ:
			    break;
		  }
       } else {
          addLoss();
          switch ( b.getMatchResultType() ) {
			  
			  case FALL:
			    addLossByFall();
				break;
			  case TECH:
			    addLossByTech();
				break;
			  case MD:
			    addLossByMD();
			    break;
			  case DEC:
			    addLossByDec();
				break;
			  case INJ:
			    if ( lossByInjury) {
			    	lossByInjuryString += ";";
			    } else {
			    	lossByInjury=true;
			    	lossByInjuryString= "Loss by Injury on ";
			    }
			    if ( b.getMatchDate() != null ) {
			    	lossByInjuryString += b.getMatchDate();
			    	lossByInjuryString += " ";
			    }
			    lossByInjuryString += " vs ";
			    lossByInjuryString += b.getOpponentTeam();
			    
			    break;
			  case DQ:
			    break;
			  case FFT:
			    break;
		  }		  
       }		  
	}
	private void addWin() { wins++; }
	private void addLoss() { losses++; }
	private void addWinByFFT() { winByFFT++; }
	private void addWinByFall() { winByFall++; }
	private void addWinByTech() { winByTech++; }
	private void addWinByMD() { winByMD++; }
	private void addWinByDec() { winByDec++; }
	private void addLossByFall() { lossByFall++; }
	private void addLossByTech() { lossByTech++; }
	private void addLossByMD() { lossByMD++; }
	private void addLossByDec() { lossByDec++; }

	public boolean equals(Object o) {
	  if ( o == null ) { return false; }
	  if ( this == o ) { return true; }
	  if ( this.getClass() == o.getClass() ){
         Wrestler ow = (Wrestler) o;		  
	     if ( this.getName().equals(ow.getName()) && this.getTeamName().equals(ow.getTeamName())) {
			 return true;
	     }
	  }
	  return false; 
	}
	public int hashCode() {
		return getName().hashCode() + getTeamName().hashCode();
	}
}
