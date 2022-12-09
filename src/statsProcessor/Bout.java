package statsProcessor;




class Bout {
	
	private String weight="";
	private String mainName="";
	private String mainTeam="";
	private String opponentName="";
	private String opponentTeam="";
	private WrestlingLanguage.WinOrLose winOrLose;
	private String matchDate="";
	private WrestlingLanguage.MatchResultType matchResultType;
	private String result="";
	private String round="";  /* this is used for tournaments. */
	private Event event;
	


	public boolean isLossByInjury() {
		if ( winOrLose == WrestlingLanguage.WinOrLose.LOSS) {
			if ( matchResultType == WrestlingLanguage.MatchResultType.INJ ) {
				return true;
			}
		}
		return false;
	}
	public String getMainName() { return mainName; }
	public String getWeight() { return weight; }
    public String getMainTeam() { return mainTeam; }
	public String getOpponentName() { return opponentName; }
	public String getOpponentTeam() { return opponentTeam; }
	public WrestlingLanguage.WinOrLose getWinOrLose() { return winOrLose; }
	public WrestlingLanguage.MatchResultType getMatchResultType() { return matchResultType; }
	public String getMatchDate() { return matchDate; }
	public String getResult() { return result; }
	public String getRound() { return round; }

	
	public Event getEvent() { return event; }
	
	public void setEvent(Event e) { event = e; }
	
	public void setMainName( String name ) {
		mainName = name;
	}
	public void setRound( String r ) {
		round = r;
	}
	
	public void setResult(String res) {
		result = res;
	}
	public void setWeight( String w ) {
		weight = w;
	}
	public void setMainTeam( String team ) {
		mainTeam = team;
	}
	public void setOpponentName( String opp ) {
		opponentName = opp;
	}
	public void setOpponentTeam( String oppTeam ) {
		opponentTeam = oppTeam;
	}

	public void setWin() {
		winOrLose = WrestlingLanguage.WinOrLose.WIN;
	}
	public void setLoss() {
		winOrLose = WrestlingLanguage.WinOrLose.LOSS;
	}
	
	public boolean isAWin() {
		if ( winOrLose ==  WrestlingLanguage.WinOrLose.WIN ) {
			return true;
		}
		return false;
	}
	public boolean isALoss() {
		if ( winOrLose == WrestlingLanguage.WinOrLose.LOSS ) {
			return true;
		}
		return false;
	}

			
	public void setMatchResultType(WrestlingLanguage.MatchResultType res) {
		matchResultType = res;
	}
    public void setMatchDate(String dt) {
		matchDate = dt;
	}
	public String toString() {
		Event e = getEvent();
		String title = e.getEventTitle();
		
		return "Weight: " + weight + " " + mainName + " " + winOrLose + " by " + result + " VS " + opponentName + " (" + opponentTeam + ") on " + matchDate + " @ " + title + " " + round;
	}
}

