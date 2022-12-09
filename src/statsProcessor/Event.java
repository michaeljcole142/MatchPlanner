/**
 * 
 */
package statsProcessor;

/**
 * @author Mike
 *
 *  This is a base class to hold common features across 
 *  dual meets and tournaments.
 */
abstract class Event {
	

	private String mainTeam;     /* this is the team that the event is for. */
	private String eventDate;    /* this is the date of the event. */
	private boolean isCompeted;  /* this is flag indicating if it has been competed or not. */
	
	private String eventTitle;
	
	public String getMainTeam() { return mainTeam; }
	public String getEventDate() { return eventDate; }
	public boolean isCompeted() { return isCompeted; }

	abstract public boolean isATournament();
	abstract public boolean isADualMeet();
	
	
	public void setEventTitle(String s) { eventTitle = s; }
	
	public String getEventTitle() { return eventTitle; }
	
	public void setMainTeam( String team ) {
		mainTeam = team;
	}
    public void setEventDate(String dt) {
		eventDate = dt;
	}
    public void setIsCompetedOn() {
		isCompeted = true;
	}
	public void setIsCompetedOff() {
		isCompeted = false;
	}
}

