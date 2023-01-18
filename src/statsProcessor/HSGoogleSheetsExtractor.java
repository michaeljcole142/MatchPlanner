package statsProcessor;
/**/

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Collections;
import java.util.List;


public class HSGoogleSheetsExtractor {
    private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "./credentials.json";

    /**
     * Creates an authorized Credential object.
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = HSGoogleSheetsExtractor.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

/*  */



/**
 * This class is used to extract information that is in track format that
 * was copied into excel.  
 * 1) results
 * 2) roster
 * 3) last year results
 */ 
 
	private static String DNP_TOKEN="DNP";
	private static String WEIGHIN_TOKEN="Weigh-ins";
	private static String DATE_TOKEN="Date";
	private static String MATCHES_TOKEN="Matches";
	private static String DASHBOARD_TOKEN="Dashboard";
	private static String TIE_BREAKING_TOKEN="Tie Breaking";
	private static String SUMMARY_TOKEN="Summary";
	private static String WEIGHTS_TOKEN="Weights";
	private static String DUAL_FOOTER_TOKEN="Team Score:";
	private static String DUAL_FOOTER_USP="USP";
	private static String FORFEIT_TOKEN="For.";
	private static String DEC_TOKEN="Dec";
	private static String SV_TOKEN="SV-1";
	private static String TB_TOKEN="TB-1";
	private static String UTB_TOKEN="UTB";
	private static String MD_TOKEN="MD";
	private static String FALL_TOKEN="Fall";
	private static String M_FOR_TOKEN="M.";
	private static String INJ_TOKEN="Inj.";
	private static String TECH_TOKEN="TF";
	private static String DQ_TOKEN="DQ";
	private static String NC_TOKEN="NC";
	private static String TEAMS_TOKEN="Teams";
	private static String DIVISIONS_TOKEN="Divisions";
	private static String WRESTLERS_TOKEN="Wrestlers";
	private static String MORE_TOKEN="More";
	private static String UNKNOWN_TOKEN="Unknown";
	private static String ROSTER_NAME_TOKEN="Name";
	private static String DUALS_TOKEN="Duals";
	private static String OFFICIAL_TOKEN="Official:";
	private static String WINNING_TEAM_TOKEN="winning team"; 
	private static String CLICK_HERE_TOKEN="Click here to get a direct"; 
	private static String GENDER_M_TOKEN="M";
	private static String GENDER_F_TOKEN="F";
	private static String GRADE_FR_TOKEN="Fr.";
	private static String GRADE_SO_TOKEN="So.";
	private static String GRADE_JR_TOKEN="Jr.";
	private static String GRADE_SR_TOKEN="Sr.";
	private static String DOUBLE_FORFEIT_TOKEN="Double Forfeit";
	private static String BYE_TOKEN="received a bye";
	private static String VS_TOKEN=" vs.";
	private static String WEIGHT_TOKEN="Weight";
	
	/*
	 * Prestige Tokens.
	 */
	private static String CH_R1_TOKEN = "Champ. Round 1";
	private static String QRTRS_TOKEN = "Quarterfinals";
	private static String SEMIS_TOKEN = "Semifinals";
	private static String FINALS_TOKEN = "1st Place Match";
	private static String CONS_3_4_TOKEN = "3rd Place Match";
	private static String CONS_R1_TOKEN = "Cons. Round 1";
	private static String CONS_SEMI_TOKEN = "Cons. Semis";
	private static String CONS_5_6_TOKEN = "5th Place Match";
	private static String CH_R2_TOKEN = "Champ. Round 2";
	private static String CONS_R4_TOKEN = "Cons. Round 4";
	private static String CONS_R5_TOKEN = "Cons. Round 5";
	private static String CONS_7_8_TOKEN = "7th Place Match";
	private static String CONS_R2_TOKEN = "Cons. Round 2";
	private static String CONS_R3_TOKEN = "Cons. Round 3";

	private static String DISTRICTS_TOKEN = "NJSIAA District";
	private static String REGIONS_TOKEN = "NJSIAA Region";
	private static String STATES_TOKEN="NJSIAA State Championships";
  
	private static String SHEET_RESULTS="Results";
	private static String SHEET_LAST_YEAR_RESULTS="LastYearResults";

	private static String SHEET_WEIGHIN_HISTORY="WeighInHistory";
	private static String SHEET_ROSTER="Roster";
	private static String SHEET_LASTYEAR_ROSTER="LastYearRoster";
	private static String SHEET_PRESTIGE_LAST_YEAR="PrestigeLastYear";
	private static String SHEET_PRESTIGE_2_YEAR="Prestige2YrsAgo";
	private static String SHEET_PRESTIGE_3_YEAR="Prestige3YrsAgo";

	
	private String sheetId;
	private String team;
    private int rowNumAt;
	private boolean verbose=true;
	private Sheets service;

	
    public HSGoogleSheetsExtractor(String t, String sid) {
		setTeam(t);
		setSheetId(sid);
	    try {
	    	final NetHttpTransport HTTP_TRANSPORT = new com.google.api.client.http.javanet.NetHttpTransport();
//	    	final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
//	    	service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
//	                .setApplicationName(APPLICATION_NAME)
//	                .build();
	    	Sheets.Builder s2 = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT));
	    	s2.setApplicationName(APPLICATION_NAME);
	        service = s2.build();
	    } catch (Exception e) { 
	    	System.out.println("ERROR->" + e.toString());
	    }
	}
    private Sheets getService() { return service; }
    
	public String getSheetId() { return sheetId; }
	public String getTeam() { return team; }
	
	public void setSheetId(String sid) { sheetId = sid; }
	public void setTeam(String t) { team = t; }
	
	private boolean isADual(String eventString) {
		String[] tokens = eventString.split(" vs. ");
		if ( tokens.length == 2 ) {
			return true;
		}
		return false;
	}
	private Bout extractDualMatch(String mString,String weight) throws ExcelExtractorException {
		int x = 0x00A0;
		char c = (char) x;
		Bout theBout = new Bout();
		theBout.setWeight(weight);
		String matchString = new String(mString);
        String cc = c  + "";
		matchString = matchString.replaceAll(cc," ");
	
		String[] tokens = matchString.split ( " over ");
		
		if ( tokens.length != 2 ) {
			if ( matchString.equals(DOUBLE_FORFEIT_TOKEN) ) {
				return null;
			}
			throw new ExcelExtractorException ("split on over failed for " + matchString,rowNumAt);
		}
		String firstString = tokens[0];
		String secondString = tokens[1];
		firstString=firstString.trim();
		secondString=secondString.trim();	
		int teamLocal = firstString.indexOf('(');
		String wrestler1 = firstString.substring(0,teamLocal-1);
        wrestler1=wrestler1.trim();
		String team1 = firstString.substring(teamLocal+1,firstString.length() - 1);
		team1=team1.trim();		
		
		String result="";
		if ( secondString.contains(") (" ) ) {
			teamLocal=secondString.indexOf(") (");
			result = secondString.substring(teamLocal+3,secondString.length() -1);
		} else {
			teamLocal= secondString.lastIndexOf('(');
			result = secondString.substring(teamLocal+1,secondString.length() -1);
		}
		
		theBout.setResult(result);
		
		String secondString2 = secondString.substring(0,teamLocal);
		secondString2=secondString2.trim();
	    String team2="";
		String wrestler2="";
		if ( result.equals(FORFEIT_TOKEN) ) {
			if ( ! secondString2.equals(UNKNOWN_TOKEN) ) {
				throw new ExcelExtractorException("Hit fft exception",rowNumAt);
		    }
			wrestler2=secondString2;
		} else {
  		  teamLocal = secondString2.indexOf('(');
		  wrestler2 = secondString2.substring(0,teamLocal-1);
		  wrestler2=wrestler2.trim();
		  team2 = secondString2.substring(teamLocal+1,secondString2.length());
		  team2=team2.trim();
		}

		String[] res = result.split(" ");

		System.out.println("res is" + result);
		result = res[0];
		
		if ( result.equals(FORFEIT_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FFT);
		} else if ( result.equals(FALL_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FALL) ;
		} else if ( result.equals(TECH_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.TECH);
		} else if ( result.equals(MD_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.MD) ;
		} else if ( result.equals(DEC_TOKEN) || result.equals(SV_TOKEN) || result.equals(TB_TOKEN) || result.equals(UTB_TOKEN) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.DEC);
		} else if ( result.equals(DQ_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.DQ);
		} else if ( result.equals(INJ_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.INJ);
		} else if ( result.equals(M_FOR_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.MFFT);
		} else {
			throw new ExcelExtractorException( "Unknown match result type encountered... <" + result + ">",rowNumAt );
		}
	
		if ( team1.equals(team) ) {
			theBout.setMainName(wrestler1);
			theBout.setMainTeam(team1);
			theBout.setWin();
			theBout.setOpponentName(wrestler2);
			theBout.setOpponentTeam(team2);
		} else if (team2.equals(team) ) {
			theBout.setMainName(wrestler2);
			theBout.setMainTeam(team2);
			theBout.setLoss();
			theBout.setOpponentName(wrestler1);
			theBout.setOpponentTeam(team1);
		} else if ( result.equals(FORFEIT_TOKEN) ) {
			return null;
		} else if ( result.contains(DOUBLE_FORFEIT_TOKEN) ) {
			return null;
		} else {
			throw new ExcelExtractorException("ERROR " + team + " not equal to " + team1 + " or " + team2,rowNumAt);
		}	
	
		return theBout;
	}
	private Bout extractTourneyMatch(String mString, String weight) throws ExcelExtractorException {
System.out.println("mString->" + mString);
		int x = 0x00A0;
		char c = (char) x;
		
		String matchString = new String(mString);
        String cc = c  + "";
		matchString = matchString.replaceAll(cc," ");
	
	    Bout theBout = new Bout();
		theBout.setWeight(weight);
		String[] tokens = matchString.split ( " over ");
		
		/* if it split on " over " it is a match.  If not, check to see if it is a bye and skip. */
		if ( tokens.length != 2 ) {
	        if ( matchString.contains(BYE_TOKEN) ) {
			  /*
               * This is a bye record.
               */
               verboseMessage("At a bye record");
              return null;
			} else if ( matchString.contains(DOUBLE_FORFEIT_TOKEN) ) {
				verboseMessage("at a Double Forfeit");
				return null;
			} else if ( matchString.contains(VS_TOKEN) ) {
				verboseMessage("at a vs match that never happened");
				return null;
            } else {
                throw new ExcelExtractorException("processing an unexpected format of tourney match record <" + mString + ">",rowNumAt);
			}			  
		}
		/* If you got here it is a match record. */
		String firstString = tokens[0];
		String secondString = tokens[1];
		firstString=firstString.trim();
		secondString=secondString.trim();
	
		/* firstString still needs the round extracted from it. */
		if ( !firstString.contains(" - ") ) {
			throw new ExcelExtractorException ("found a tourney record without a round on it.  <" + mString + ">",rowNumAt);
		} 
		int i = firstString.indexOf(" - ");

		String roundString = firstString.substring(0,i);
	    theBout.setRound(roundString);		
		firstString=firstString.substring(i+2);
	 
		int teamLocal = firstString.indexOf('(');
		String wrestler1 = firstString.substring(0,teamLocal-1);
        wrestler1=wrestler1.trim();
		String team1 = firstString.substring(teamLocal+1,firstString.length() - 1);
		team1=team1.trim();		
		
		teamLocal = secondString.lastIndexOf('(');
		String result = secondString.substring(teamLocal+1,secondString.length() -1);
		
		theBout.setResult(result);
		
		String secondString2 = secondString.substring(0,teamLocal);
		secondString2=secondString2.trim();
		
	    String team2="";
		String wrestler2="";
		String[] res = result.split(" ");
		if ( result.equals(FORFEIT_TOKEN) ) {
			verboseMessage("hit fft in tourney ");
			/* don't mess with result. */
		} else {
			result=res[0];
		}
		
  		teamLocal = secondString2.indexOf('(');

        if ( teamLocal < 0 && secondString2.equals(UNKNOWN_TOKEN) && result.equals(FORFEIT_TOKEN) ) {
			/* there is no second wrestler. */
			wrestler2="";			//wrestler2 = secondString2.substring(0,teamLocal-1);
			team2="";
		} else {
			wrestler2 = secondString2.substring(0,teamLocal-1);
			wrestler2=wrestler2.trim();
			team2 = secondString2.substring(teamLocal+1,secondString2.length()-1);
			team2=team2.trim();
		}
        if ( result.equals("Fall)") ) {
        	System.out.println("here");
        }
		if ( result.equals(FORFEIT_TOKEN) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FFT);
		} else if ( result.equals(M_FOR_TOKEN) ) {
			if ( res.length > 1 & res[1].equals(FORFEIT_TOKEN) ) {
				theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FFT);
			}
		} else if ( result.equals(FALL_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FALL) ;
		} else if ( result.equals(TECH_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.TECH);
		} else if ( result.equals(MD_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.MD) ;
		} else if ( result.equals(DEC_TOKEN)|| result.equals(SV_TOKEN) || result.equals(TB_TOKEN) || result.equals(UTB_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.DEC);
		} else if ( result.equals(DQ_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.DQ);
		} else if ( result.equals(INJ_TOKEN)) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.INJ);
		} else if ( result.equals(NC_TOKEN )) {	
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.NC);
		} else {
			throw new ExcelExtractorException( "Unknown match result type encountered... <" + result + ">" ,rowNumAt);
		}
		
		if ( team1.equals(team) ) {
			theBout.setMainName(wrestler1);
			theBout.setMainTeam(team1);
			theBout.setWin();
			theBout.setOpponentName(wrestler2);
			theBout.setOpponentTeam(team2);
			
		} else if (team2.equals(team) ) {
			theBout.setMainName(wrestler2);
			theBout.setMainTeam(team2);
			theBout.setLoss();
			theBout.setOpponentName(wrestler1);
			theBout.setOpponentTeam(team1);
		} else {
			throw new ExcelExtractorException("ERROR " + team + " not equal to " + team1 + " or " + team2,rowNumAt);
		}	
		
		return theBout;
		
	}
	private DualMeet initializeDual(String eventString) throws ExcelExtractorException {
		String[] tokens = eventString.split(" vs. ");
		if (tokens.length != 2 ) {
			throw new ExcelExtractorException ("error in initializeDualString with " + eventString + " length is " + tokens.length,rowNumAt);
		}
		DualMeet theDual = new DualMeet();
		theDual.setMainTeam(team);
		int lastParen = tokens[1].lastIndexOf('(');
		String team1 = tokens[0];
		String team2 = tokens[1].substring(0,lastParen -1);
		String dt = tokens[1].substring(lastParen+1,tokens[1].length()-1);
		if ( team1.equals(team) ) { 
			theDual.setOpponent(team2);
		} else {
			theDual.setOpponent(team1);
		}
		theDual.setEventDate(dt);
		theDual.setEventTitle(team2 + " Dual");
		
		return theDual;
	}
	private Tournament initializeTourney(String eventString) throws ExcelExtractorException {
		
		Tournament theTourney = new Tournament();
		theTourney.setMainTeam(team);
		theTourney.setEventTitle(eventString);
		return theTourney;
	}
		
	private boolean isATourney(String eventString) {
	  /*
	   * For now, if it is not a dual, it must be a tourney.
	   */
	   if ( isADual(eventString) ) {
		   return false;
	   }
	   return true;
	}
	private boolean isMatchHeaderToken(List<Object> thisRow ) {
		if ( thisRow.size() >= 1 ) {
			String token = thisRow.get(0).toString();
			if ( token.equals(MATCHES_TOKEN) || 
				token.equals(DASHBOARD_TOKEN) || 
				token.equals(TIE_BREAKING_TOKEN) || 
				token.equals(SUMMARY_TOKEN) || 
				token.equals(DUALS_TOKEN) || 
				token.equals(MORE_TOKEN) || 
				token.equals(WEIGHTS_TOKEN) ) {
				verboseMessage("Token is match header " + token); 
				return true;
			}
		}
		verboseMessage("Token is not a match header "); 
		return false;
	}
	private void processDualHeaderRows(List<List<Object>> resultSheet) throws Exception {
		String checkCell;
		boolean stillProcessing = true;
		while ( stillProcessing ) {
			System.out.println("processing->" + rowNumAt);
			
			List<Object> nextRow = resultSheet.get(rowNumAt);
			if (nextRow.size()==0) {
				verboseMessage("empty row");
				rowNumAt++;
			} else {
	System.out.println("in else");
				checkCell = nextRow.get(0).toString();
				if ( isABlankRow(nextRow)) {
					verboseMessage("skipping blank row");
					rowNumAt++;
				} else {
	System.out.println("in else else->" + checkCell + " size->" + nextRow.size() );
					if ( checkCell != null && checkCell.length() > 0 ) {
						System.out.println("not null");
						if ( checkCell.equals(MATCHES_TOKEN) ||
							checkCell.equals(DASHBOARD_TOKEN) ||
							checkCell.equals(TIE_BREAKING_TOKEN) ||
							checkCell.equals(SUMMARY_TOKEN) ||
							checkCell.equals(WEIGHTS_TOKEN) ||
							checkCell.equals(DUALS_TOKEN) ||
							checkCell.equals(MORE_TOKEN) ) {
							verboseMessage("At Dual header record" ) ;
							rowNumAt++;
						}
					} else {
		System.out.println("in else else else");
						String wCell = nextRow.get(1).toString();
						if (wCell != null ) {
							if (wCell.equals(WEIGHT_TOKEN) ) {
								verboseMessage("Past Dual Headers");
								stillProcessing = false;
								rowNumAt--;
							}
						} else {
							throw new ExcelExtractorException("Unknown Dual Header",rowNumAt);	
						}
					}
				}
				
			}
		   
	   }
	 
	   /* Now should find dual header row. */
	   rowNumAt++;
	   List<Object> dualHeaderRow = resultSheet.get(rowNumAt);
	   if ( dualHeaderRow == null ) {
		   throw new ExcelExtractorException("empty row when expected dual header row at " + rowNumAt,rowNumAt);
	   }
	   int lastCellNum = dualHeaderRow.size()-1;
	   if ( lastCellNum != 4 ) {
		   throw new ExcelExtractorException ("Unexpected Header Count of " + lastCellNum + " at " + rowNumAt,rowNumAt);
	   }
	   
	   verboseMessage("Processed Dual Meet Headers Successfully.");
	   return;
  }	
	private int processTourneyHeaderRows(List<List<Object>> resultSheet) throws ExcelExtractorException {
       int rowCheck=1;
	   String checkCell;
	   List<Object> teamsRow = resultSheet.get(rowNumAt+rowCheck);
	   if (teamsRow == null) {
		   throw new ExcelExtractorException("Didn't find Teams record. ",rowNumAt);
	   }
	   checkCell = teamsRow.get(0).toString();
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Teams row was empty. ",rowNumAt);
	   }
	   if (! checkCell.equals(TEAMS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted non Teams record found.  " + checkCell, rowNumAt );
       }
	   rowCheck++;
       List<Object> divisionsRow = resultSheet.get(rowNumAt+rowCheck);
	   if (divisionsRow == null) {
		   throw new ExcelExtractorException("Didn't find Divisions record. ",rowNumAt);
	   }
	   checkCell = divisionsRow.get(0).toString();
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Divisions row was empty. ",rowNumAt);
	   }
	   if (!checkCell.equals(DIVISIONS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Tourney record found.  " + checkCell,rowNumAt );
       }
	   rowCheck++;
       List<Object> weightsRow = resultSheet.get(rowNumAt+rowCheck);
	   if (weightsRow == null) {
		   throw new ExcelExtractorException("Didn't find Weights record. ",rowNumAt);
	   }
	   checkCell = weightsRow.get(0).toString();
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Weights row was empty. ",rowNumAt);
	   }
	   if (!checkCell.equals(WEIGHTS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Weights record found.  " + checkCell,rowNumAt );
       }
	   rowCheck++;
       List<Object> wrestlersRow = resultSheet.get(rowNumAt+rowCheck);
	   if (wrestlersRow == null) {
		   throw new ExcelExtractorException("Didn't find Wrestlers record. ",rowNumAt);
	   }
	   checkCell = wrestlersRow.get(0).toString();
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Wrestlers row was empty. ",rowNumAt);
	   }
	   if (!checkCell.equals(WRESTLERS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Dual record found.  " + checkCell,rowNumAt );
       }
	   rowCheck++;
       List<Object> moreRow = resultSheet.get(rowNumAt+rowCheck);
	   if (moreRow == null) {
		   throw new ExcelExtractorException("Didn't find More record. ",rowNumAt);
	   }
	   checkCell = moreRow.get(0).toString();
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("More row was empty. ",rowNumAt);
	   }
	   if (!checkCell.equals(MORE_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Dual record found.  " + checkCell,rowNumAt );
       }
	   /* Should find a blank row next. */
	   rowCheck++;
	   List<Object> blankRow = resultSheet.get(rowNumAt+rowCheck);

	   if ( ! isABlankRow(blankRow) ) {
		   checkCell = blankRow.get(0).toString();
		   if ( checkCell.substring(0,5).equals("Team:") ) {
			  rowCheck++;
			  blankRow = resultSheet.get(rowNumAt+rowCheck);
			  if ( !isABlankRow(blankRow)) {
				  throw new ExcelExtractorException ("expecting blank row  or Team: at " + (rowNumAt+rowCheck),rowNumAt);
			  }
		   } else { 
			   throw new ExcelExtractorException ("expecting blank row  or Team: at " + (rowNumAt+rowCheck),rowNumAt);
		   }
	   }
	   /* Now should find tourney header row. */
	   rowCheck++;
	   List<Object> tourneyHeaderRow = resultSheet.get(rowNumAt+rowCheck);
	   if ( tourneyHeaderRow == null ) {
		   throw new ExcelExtractorException("empty row when expected tourney header row at " + (rowNumAt+rowCheck),rowNumAt);
	   }
	   int lastCellNum = tourneyHeaderRow.size();
	   if ( lastCellNum < 5 ) {
		   throw new ExcelExtractorException ("Unexpected Header Count of " + lastCellNum + " at " + (rowNumAt+rowCheck),rowNumAt);
	   }
	   
	   verboseMessage("Processed Tourney Headers Successfully."); 
	   return rowCheck;
  }	
	private int processDualMatches(DualMeet d, List<List<Object>> resultSheet) throws ExcelExtractorException {
       int rowCheck=1;
	   /*
	    *   Should find 14 weights in a dual.  
		*/
       // NEED TO FIX WEIGHT CLASS NUMBERS
	   for ( int i=0; i<14; i++ ) {
			List<Object> aRow = resultSheet.get(rowNumAt+rowCheck);
			String weightCell = aRow.get(1).toString();
			String matchCell = aRow.get(2).toString();
			Bout b = extractDualMatch(matchCell,weightCell);
			if ( b == null ) {
				verboseMessage("Processed non match row..." );
			} else {
				b.setMatchDate(d.getEventDate());
				if ( b.getOpponentTeam() == null || b.getOpponentTeam().length() == 0 ) {
					b.setOpponentTeam(d.getOpponent());
				}
				d.addBout(b);	
				b.setEvent(d);
				verboseMessage("Processed...." + b ); 
			}
			rowCheck++;
	   }
       if ( d.isFullDual() ) {
 	       verboseMessage("Processed Dual Meet  Successfully."); 
	   } else {
		   verboseMessage("This is not a full dual meet." ) ; 
       }
	   return rowCheck;
    }	
	private int processTourneyMatches(Tournament t, List<List<Object>> resultSheet) throws ExcelExtractorException {
       int rowCheck=1;
	   /*
	    *   Process until a blank row.  
		*/
		boolean atBlankRow = false;
		
	   while ( ! atBlankRow && rowNumAt+rowCheck < resultSheet.size()) {
		   List<Object> aRow = resultSheet.get(rowNumAt+rowCheck);
		   if ( aRow == null || aRow.size()==0 ) {
			   atBlankRow = true;
			   rowCheck--;
		   } else if ( aRow.size() == 1 ) {
			   atBlankRow = true;
			   rowCheck--; 
		   } else {
				String weightCell = aRow.get(1).toString();
				String matchCell = aRow.get(2).toString();
				
				System.out.println("---->"  + weightCell + " -->" + matchCell);
				Bout b = extractTourneyMatch(matchCell,weightCell);
				/* if the bout returned is null that means it was a bye round and we skip it. */
				if ( b != null ) {   
			     b.setMatchDate(t.getEventDate());
				 t.addBout(b);
				 b.setEvent(t);
			     verboseMessage("Processed...." + b );
				}
				rowCheck++;
			}
	   }

	   verboseMessage("Processed Tourney  Successfully.");
	   return rowCheck;
    }	
    private void verboseMessage(String message) {
		if ( verbose ) {
			System.out.println("ROWAT<" + (rowNumAt+1) + ">-" + message);
		}
	}
	private Team extractRoster(Team t) throws Exception {

		rowNumAt=0;
		
		String range = SHEET_ROSTER + "!A:G";
		
		ValueRange response = service.spreadsheets().values()
	                .get(this.getSheetId(), range)
	                .execute();
	    List<List<Object>> values = response.getValues();
		
	    verboseMessage("Roster: " + values.size());;
		/*
		 * This will read in an initialize the team roster.
		 * 
		 */
		while ( rowNumAt < values.size() ) {			  
			verboseMessage("ROSTER: PROCESSING ROW " + rowNumAt + "..." );
			List<Object> rowAt = values.get(rowNumAt);
			if ( rowAt ==  null ) {
				/* Skip null rows */
				verboseMessage("ROSTER: null row. ");  
			} else {
				verboseMessage("ROSTER: ROW<" + rowNumAt + "> Cells<" + rowAt.size() + ">");
				if ( rowAt.size() >= 7 ) {
					String nameCell = rowAt.get(1).toString();
					if ( nameCell.equals(ROSTER_NAME_TOKEN) ) {
						verboseMessage("ROSTER: at header row.");
					} else {
					
						String wtClassVal = rowAt.get(3).toString();
						String genderVal = rowAt.get(4).toString();
						String gradeVal = rowAt.get(5).toString();
						String recordVal = rowAt.get(6).toString();
						verboseMessage( nameCell + wtClassVal + genderVal + gradeVal + recordVal );
						if ( t.wrestlerExists(nameCell) ) {
							throw new ExcelExtractorException("LAST YEAR ROSTER: wrestler " + nameCell + " exists already.",rowNumAt);
						} else {
							Wrestler aWrestler = new Wrestler(nameCell,team);
							aWrestler.setTrackWtClass(wtClassVal);
							aWrestler.setGender(getGender(genderVal));
							aWrestler.setGrade(getGrade(gradeVal));
							aWrestler.setTrackRecord(recordVal);
							t.addWrestler(aWrestler);
							aWrestler.printVerbose();
						}
					}
				} else {
					verboseMessage("ROSTER: Skipping Row...");
				}
			} 
			rowNumAt++;
		}
		return t;
	}
	private Team extractWeighInHistory(Team t) throws Exception {
		rowNumAt=0;
		
		String range = SHEET_WEIGHIN_HISTORY + "!A:E";
		
		ValueRange response = service.spreadsheets().values()
	                .get(this.getSheetId(), range)
	                .execute();
	    List<List<Object>> values = response.getValues();
		
		if ( values == null || values.size() == 0 ) {
			return t;
		}

		verboseMessage("Weigh In History: rows->" + values.size());
		/*
		 * This will read in and Build a history then add it to a wrestler.
		 */
		
		
		Wrestler wrestlerAt;
		while ( rowNumAt < values.size() ) {			  
			verboseMessage("WeighInHistory: PROCESSING ROWS ...");
			List<Object> rowAt = values.get(rowNumAt);
			if ( rowAt ==  null ) {
				/* Skip null rows */
				verboseMessage("WeighInHistory: null row. ");  
			} else {
				verboseMessage("WeighInHistory: Cells<" + rowAt.size() + ">");
				String header = rowAt.get(0).toString();
			
				verboseMessage("Wrestler " + header );
				String wrestlerName = header.substring(0,header.indexOf(WEIGHIN_TOKEN)-1);
				wrestlerAt=t.getWrestler(wrestlerName);
				WeighInHistory wiHistory = new WeighInHistory();
				wiHistory.setName(wrestlerName);
				if ( wrestlerAt == null ) {
					throw new ExcelExtractorException("Can not find wrestler " + wrestlerName, rowNumAt );
				}
				verboseMessage("WeighInHistory: <" + header + ">");
				rowNumAt++;
				boolean pHeaders = true;
				while ( pHeaders ) {
					rowAt=values.get(rowNumAt);
					if ( rowAt.size() <= 1 ) {
						rowNumAt++;
					} else {
						String firstVal = rowAt.get(0).toString();
						String secondVal = rowAt.get(1).toString();
						if ( firstVal != null && firstVal.length() > 0 ) {
							/* still at headers */
							rowNumAt++;
						} else if ( secondVal != null && secondVal.length() > 0 ) {
							/* now on a wi row. */
							pHeaders = false;
						} else if ( isABlankRow(rowAt) ) {
							verboseMessage("Skipping Blank Row");
							rowNumAt++;
						} else {
							throw new ExcelExtractorException("WeighInHistory: At exception row. ", rowNumAt );
						}
					}
				}
				verboseMessage("WeighInHistory: Finished Header rows. ");
				/* processing weigh in records header now */
				String dt = rowAt.get(1).toString();
				if ( dt == null ) { 
					throw new ExcelExtractorException("Unexpected empty cell at row <",rowNumAt);
				}
				if ( ! dt.equals(DATE_TOKEN) ) {
					throw new ExcelExtractorException("Unexpected token <" + dt + ">", rowNumAt);
				}
				rowNumAt++;
				boolean pRow = true;
				while ( pRow && rowNumAt < values.size()) {
					rowAt=values.get(rowNumAt);
					System.out.println("got->" + rowNumAt + " size->" + rowAt.size());
					if (  isABlankRow(rowAt) ) {
						/* skip blank row. */
						verboseMessage("skip blank row");
						rowNumAt++;
					} else {
						String ckCell = rowAt.get(0).toString();
						if ( ckCell == null || ckCell.length() == 0 ) {
							WeighIn wi = new WeighIn();
							wi.setWIDate(rowAt.get(1).toString());
							wi.setWIEvent(rowAt.get(2).toString());
							
							String wiString = rowAt.get(3).toString();
							if ( wiString != null ) {
								if ( wiString.equals(DNP_TOKEN) ) {
									wi.setWIWeight(0);
								} else {
									Float wt = Float.valueOf(wiString);
									wi.setWIWeight(wt);
								}
								if ( rowAt.size() > 4 ) {
									String ss = rowAt.get(4).toString();
									if ( ss != null ) {
										wi.setEntryDateTime(ss);
									}
								}
								wiHistory.addWI(wi);
							}
							rowNumAt++;
						} else {
							System.out.println("ckCell = " + ckCell );
							pRow = false;
							rowNumAt--;
						}
					}
					
				}
				wrestlerAt.setWeighInHistory(wiHistory);
			} 
			rowNumAt++;
		}
		return t;
	}
	private WrestlingLanguage.Gender getGender(String g )  throws ExcelExtractorException {
		if ( g.equals( GENDER_M_TOKEN ) ) {
			return WrestlingLanguage.Gender.M;
		} else if ( g.equals( GENDER_F_TOKEN ) ) {
			return WrestlingLanguage.Gender.F;
		} else {
			throw new ExcelExtractorException("Unknown Gender <" + g + "> found." ,rowNumAt);
		}
	}
    private WrestlingLanguage.Grade getGrade(String gs ) throws ExcelExtractorException {
		if ( gs.equals ( GRADE_FR_TOKEN ) ) {
			return WrestlingLanguage.Grade.FR;
		} else if ( gs.equals ( GRADE_SO_TOKEN ) ) {
			return WrestlingLanguage.Grade.SO;
	    } else if ( gs.equals ( GRADE_JR_TOKEN ) ) {
			return WrestlingLanguage.Grade.JR;
		} else if ( gs.equals ( GRADE_SR_TOKEN ) ) {
			return WrestlingLanguage.Grade.SR; 
		} else {
			throw new ExcelExtractorException( "Unknown Grade Token found <" + gs + ">", rowNumAt);
		}
	}
	private Team extractLastYearRoster(Team t) throws Exception {
		
		rowNumAt=0;
		
		String range = SHEET_LASTYEAR_ROSTER + "!A:G";
		
		ValueRange response = service.spreadsheets().values()
	                .get(this.getSheetId(), range)
	                .execute();
	    List<List<Object>> values = response.getValues();
		
	    verboseMessage("Last Year Roster: " + values.size());;
		/*
		 * This will read in an initialize the team roster.
		 * 
		 */
		while ( rowNumAt < values.size() ) {			  
			verboseMessage("LAST YEAR ROSTER: PROCESSING ROW " + rowNumAt + "..." );
			List<Object> rowAt = values.get(rowNumAt);
			if ( rowAt ==  null ) {
				/* Skip null rows */
				verboseMessage("LAST YEAR ROSTER: null row. ");  
			} else {
				verboseMessage("LAST YEAR ROSTER: ROW<" + rowNumAt + "> Cells<" + rowAt.size() + ">");
				if ( rowAt.size() >= 7 ) {
					String nameCell = rowAt.get(1).toString();
					if ( nameCell.equals(ROSTER_NAME_TOKEN) ) {
						verboseMessage("LAST YEAR ROSTER: at header row.");
					} else {
					
						String wtClassVal = rowAt.get(3).toString();
						String genderVal = rowAt.get(4).toString();
						String gradeVal = rowAt.get(5).toString();
						String recordVal = rowAt.get(6).toString();
						verboseMessage( nameCell + wtClassVal + genderVal + gradeVal + recordVal );
						if ( t.wrestlerLastYearExists(nameCell) ) {
							throw new ExcelExtractorException("LAST YEAR ROSTER: wrestler " + nameCell + " exists already.",rowNumAt);
						} else {
							Wrestler aWrestler = new Wrestler(nameCell,team);
							aWrestler.setTrackWtClass(wtClassVal);
							aWrestler.setGender(getGender(genderVal));
							aWrestler.setGrade(getGrade(gradeVal));
							aWrestler.setTrackRecord(recordVal);
							t.addLastYearWrestler(aWrestler);
							aWrestler.printVerbose();
						}
					}
				} else {
					verboseMessage("LAST YEAR ROSTER: Skipping Row...");
				}
			} 
			rowNumAt++;
		}
		return t;
	}
	private boolean firstColOnlyContent(List<Object> r) {
		if ( r.size() == 1) { return true; }

		return false;
	}
	/*
	 * This is used to skip rows at the footer of a dual that are probably
	 * points for misconduct. 
	 */
	private boolean garbageRecord(List<Object> r) {

		String matchCell = r.get(2).toString();
		if ( matchCell != null ) {
			if ( matchCell.length() == 0 ) {
				return true;
			}
		} else {
			return true;
		}
		return false;
	}	
		
	private boolean isABlankRow(List<Object> r) {
		if ( r == null ) { return true; }
		if ( r.size() == 0 ) { return true; }
		if ( r.size() == 1 && r.get(0).toString().length() == 0 ) { return true; }
		return false;
	}
    public Team extractResults(Team theTeam) throws Exception {
		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
		rowNumAt=0;
		
		String range = SHEET_RESULTS + "!A:E";
		
		ValueRange response = service.spreadsheets().values()
	                .get(this.getSheetId(), range)
	                .execute();
	    List<List<Object>> values = response.getValues();
		
		if ( values == null || values.size() == 0 ) {
			return theTeam;
		}

		/*
		 * This while loop will process chunks of the results.  This is why the rowAt is not incremented
		 * in the main line of a for loop.
		 */
		boolean inAnEventFlag = false;   /* this flag tells me if we are processing an event */
		while ( rowNumAt < values.size() ) {			  
			verboseMessage("PROCESSING ROW ...");
			  
			List<Object> rowAt = values.get(rowNumAt);
			if ( rowAt.size() == 0 ) {
				/* Skip empty rows */
				verboseMessage(" empty row. ");
				rowNumAt++;
			} else {
				/* Check and See if we are at a new event. */
				if ( inAnEventFlag == false ) {
                    verboseMessage("InAnEventFlag is false"); 
					if ( ! firstColOnlyContent(rowAt) ) {
						if ( garbageRecord(rowAt) ) {
							verboseMessage("Skipping a garbage record");
						} else {
							throw new ExcelExtractorException("Something weird PhysicalNumberOfCells is <" + rowAt.size() + "> expecting 0",rowNumAt);
						}	
					} else {
						String cellZero = rowAt.get(0).toString();
						if ( isMatchHeaderToken(rowAt) ) {
							throw new ExcelExtractorException("Something weird",rowNumAt);
						} else if ( cellZero.startsWith(OFFICIAL_TOKEN) ) {
							verboseMessage("Skip row, it is a lingering Official row from dual.");
						} else if ( cellZero.startsWith(WINNING_TEAM_TOKEN) ) {
							verboseMessage("Skip row, winning team footer.");
						} else if ( cellZero.startsWith(CLICK_HERE_TOKEN) ) {
							verboseMessage("Skip row, click here row found.");
						} else {
						  /* 
						   * We are at a new event.  Next step is to see if it is a tournament or a dual.
						   */
						  String eventString = cellZero;
					      verboseMessage(" is a match <" + eventString + ">");
						  inAnEventFlag = true;
						  /*
						   *  Check to see if it is a dual
						   */
						   if ( this.isADual(eventString) ) {
							   verboseMessage(" and dual confirmed. ");
							   DualMeet d = this.initializeDual(eventString);
							   rowNumAt++;
							   /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								*/
							   this.processDualHeaderRows(values);
							   
							   verboseMessage("----- Dual:" + d );
							   /*
							    * Now we process the dual itself.
							    */
							   int addRow = this.processDualMatches(d,values);
							   rowNumAt += addRow;
							   verboseMessage("rowNumAt now " + rowNumAt);
							   /* look for trailer record. */
							   rowAt = values.get(rowNumAt);
							   if (rowAt != null) {
								   String c = rowAt.get(0).toString();
								   if ( c.equals(DUAL_FOOTER_USP)) {
									  verboseMessage("At Unsportsmanlike point");
									  rowNumAt++;
									  rowAt=values.get(rowNumAt);
									  c = rowAt.get(0).toString();

								   }
								   if ( c.equals(DUAL_FOOTER_TOKEN) ) {
									   verboseMessage("At Dual Footer row " );
								   } else {
									   rowNumAt--;
								   }
							   }
							   theTeam.addDualMeet(d);
							   inAnEventFlag=false;
						   } 
						   /*
						    * If it is not a dual, check to see if it is a tournament.
						    */
						    else if ( isATourney(eventString) ) {
								Tournament t = initializeTourney(eventString);
								verboseMessage(" and a tourney confirmed. ");
								 /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								*/
							   int addRow = this.processTourneyHeaderRows(values);
							   rowNumAt += addRow; 
							   verboseMessage("----- Tourney:" + t);
							    /*
							    * Now we process the tourney itself.
							    */
							   addRow = this.processTourneyMatches(t,values);
							   rowNumAt += addRow;
							   verboseMessage("rowNumAt now " + rowNumAt);
							   theTeam.addTournament(t);
							   inAnEventFlag=false;
						    }
						   /* 
						    * Something went wrong, throw an exception....
							*/
							else {
								throw new ExcelExtractorException("Can't determine if dual or tourney. ",rowNumAt);
						    }
						}
					}	
					rowNumAt++;
				} else {   /*here we are already in an event */
					if ( isMatchHeaderToken(rowAt) ){
						verboseMessage(" is a match token") ;
					} else {
						verboseMessage(" is not a match token");
				    }
					rowNumAt++;
				}
			}
		}
		return theTeam;
	}
	private void setPrestige(Wrestler w, String tourney, String round, WrestlingLanguage.WinOrLose worl, int year ) throws ExcelExtractorException {

		WrestlingLanguage.Prestige pNow = null;
		
		if ( tourney.equals(DISTRICTS_TOKEN) ) {			
			if ( round.equals(CH_R1_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.DQuarters;
				} else {
					pNow = WrestlingLanguage.Prestige.DPrelim;
				}
			} else if ( round.equals(CH_R2_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(QRTRS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.D4th;
				} else {
					pNow = WrestlingLanguage.Prestige.DQuarters;
				}
			} else if ( round.equals(SEMIS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.R2v3;
				} else {
					pNow = WrestlingLanguage.Prestige.D4th;
				}
			} else if ( round.equals(FINALS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.RWB1;
				} else {
					pNow = WrestlingLanguage.Prestige.R2v3;
				}
			} else if ( round.equals(CONS_3_4_TOKEN) ){
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.R2v3;
				} else {
					pNow = WrestlingLanguage.Prestige.D4th;
				}
			} else if ( round.equals(CONS_5_6_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_7_8_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R1_TOKEN) ){
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R2_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R3_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R4_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R5_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_SEMI_TOKEN) ){
				throw new ExcelExtractorException("Unexpected Districts round token <" + round + ">", rowNumAt);
			} else {
				throw new ExcelExtractorException("Unexpected round token found <" + round + ">", rowNumAt);
			}

		} else if ( tourney.equals(REGIONS_TOKEN) ) {
			if ( round.equals(CH_R1_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.RWB1;
				} else {
					pNow = WrestlingLanguage.Prestige.R2v3;
				}
			} else if ( round.equals(CH_R2_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Regions round token <" + round + ">", rowNumAt);
			} else if ( round.equals(QRTRS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.R6th;
				} else {
					pNow = WrestlingLanguage.Prestige.RWB1;
				}
			} else if ( round.equals(SEMIS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB1;
				} else {
					pNow = WrestlingLanguage.Prestige.D4th;
				}
			} else if ( round.equals(FINALS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB1;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB1;
				}
			} else if ( round.equals(CONS_3_4_TOKEN) ){
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB1;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB1;
				}
			} else if ( round.equals(CONS_5_6_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.R5th;
				} else {
					pNow = WrestlingLanguage.Prestige.R6th;
				}
			} else if ( round.equals(CONS_7_8_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Regions round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R1_TOKEN) ){
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.R6th;
				} else {
					pNow = WrestlingLanguage.Prestige.RWB1;
				}
			} else if ( round.equals(CONS_R2_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Regions round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R3_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Regions round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R4_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Regions round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_R5_TOKEN) ) {
				throw new ExcelExtractorException("Unexpected Regions round token <" + round + ">", rowNumAt);
			} else if ( round.equals(CONS_SEMI_TOKEN) ){
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB1;
				} else {
					pNow = WrestlingLanguage.Prestige.R6th;
				}
			} else {
				throw new ExcelExtractorException("Unexpected round token found <" + round + ">", rowNumAt);
			}
		} else if ( tourney.equals(STATES_TOKEN) ) {
			if ( round.equals(CH_R1_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB2;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB1;
				}
			} else if ( round.equals(CH_R2_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB3;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB2;
				}
			} else if ( round.equals(QRTRS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S6th;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB3;
				}
			} else if ( round.equals(SEMIS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S2nd;
				} else {
					pNow = WrestlingLanguage.Prestige.S6th;
				}
			} else if ( round.equals(FINALS_TOKEN)) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S1st;
				} else {
					pNow = WrestlingLanguage.Prestige.S2nd;
				}
			} else if ( round.equals(CONS_3_4_TOKEN) ){
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S3rd;
				} else {
					pNow = WrestlingLanguage.Prestige.S4th;
				}
			} else if ( round.equals(CONS_5_6_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S5th;
				} else {
					pNow = WrestlingLanguage.Prestige.S6th;
				}
			} else if ( round.equals(CONS_7_8_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S7th;
				} else {
					pNow = WrestlingLanguage.Prestige.S8th;
				}
			} else if ( round.equals(CONS_R1_TOKEN) ){
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB2;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB1;
				}
			} else if ( round.equals(CONS_R2_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB3;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB2;
				}
			} else if ( round.equals(CONS_R3_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB3;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB3;
				}
			} else if ( round.equals(CONS_R4_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.SWB3;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB3;
				}
			} else if ( round.equals(CONS_R5_TOKEN) ) {
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S8th;
				} else {
					pNow = WrestlingLanguage.Prestige.SWB3;
				}
			} else if ( round.equals(CONS_SEMI_TOKEN) ){
				if ( worl == WrestlingLanguage.WinOrLose.WIN ) {
					pNow = WrestlingLanguage.Prestige.S4th;
				} else {
					pNow = WrestlingLanguage.Prestige.R6th;
				}
			} else {
				throw new ExcelExtractorException("Unexpected round token found <" + round + ">", rowNumAt);
			}

		} else {
			throw new ExcelExtractorException("Unexpected Prestige tourney found <" + tourney, rowNumAt);

		}
		if ( pNow != null  ) {
			if ( year == 1 ) {
				if ( w.getPrestigeLastYear() == null  ) {
					w.setPrestigeLastYear(pNow);
				} else {
					if ( pNow.ordinal() > w.getPrestigeLastYear().ordinal() ) {
						w.setPrestigeLastYear(pNow);
					}
				}
			} else if ( year == 2 ) {
				if ( w.getPrestige2YearsAgo() == null  ) {
					w.setPrestige2YearsAgo(pNow);
				} else {
					if ( pNow.ordinal() > w.getPrestige2YearsAgo().ordinal() ) {
						w.setPrestige2YearsAgo(pNow);
					}
				}
			} else if ( year == 3 ) {
				if ( w.getPrestige3YearsAgo() == null  ) {
					w.setPrestige3YearsAgo(pNow);
				} else {
					if ( pNow.ordinal() > w.getPrestige3YearsAgo().ordinal() ) {
						w.setPrestige3YearsAgo(pNow);
					}
				}
			} else {
				throw new ExcelExtractorException("Unexpected Year passed " + year, rowNumAt);
			}
		}

	}
	private void syncPrestige(Team theTeam, int year) throws ExcelExtractorException {
	
		
		List<Tournament> lyt;
		if ( year == 1 ) {
			lyt = theTeam.getPrestigeLastYear();
		} else if ( year == 2 ) {
			lyt = theTeam.getPrestige2YearsAgo();
		} else if ( year == 3 ) {
			lyt = theTeam.getPrestige3YearsAgo();			
		} else {
			throw new ExcelExtractorException("Unexpected year " + year, rowNumAt);
		}
		
		for ( int i=0; i < lyt.size(); i++ ) {
			Tournament t = lyt.get(i);
			List<Bout> theBouts = t.getBouts();
			for ( int ii=0; ii < theBouts.size(); ii++ ) {
				String round = theBouts.get(ii).getRound();
				String name = theBouts.get(ii).getMainName();
				String tourney = "";
				if ( t.getEventTitle().contains(DISTRICTS_TOKEN) ) {
					tourney=DISTRICTS_TOKEN;
				} else if ( t.getEventTitle().contains(REGIONS_TOKEN) ) {
					tourney=REGIONS_TOKEN;
				} else if ( t.getEventTitle().contains(STATES_TOKEN) ) {
					tourney=STATES_TOKEN;
				} else {
					throw new ExcelExtractorException("Unexexpected Event string <" + t.getEventTitle(), rowNumAt);
				}
				Wrestler w = theTeam.getWrestler(name);
				if ( w != null ) {
					setPrestige(w,tourney,round,theBouts.get(ii).getWinOrLose(),year);
				}
			}
		}
		return;
	}
	public Team extractPrestige(Team theTeam,int year) throws Exception {
		/*
		 * This is the processing of last year's prestige sheet.
		 */	
		rowNumAt=0;
		String range="";
		if ( year == 1 ) {
		  range = SHEET_PRESTIGE_LAST_YEAR + "!A:E";
		} else if ( year == 2) {
			range = SHEET_PRESTIGE_2_YEAR + "!A:E";
		} else if ( year == 3) {
			range = SHEET_PRESTIGE_3_YEAR + "!A:E";
		} else { 
			throw new ExcelExtractorException("Prestige Year unknown <" + year + "> expecting 1,2 or 3",rowNumAt);

		}
		ValueRange response = service.spreadsheets().values()
	                .get(this.getSheetId(), range)
	                .execute();
	    List<List<Object>> values = response.getValues();
		
		if ( values == null  || values.size() == 0 ) {
			return theTeam;
		}
	
		/*
		 * This while loop will process chunks of the results.  This is why the rowAt is not incremented
		 * in the main line of a for loop.
		 */
		boolean inAnEventFlag = false;   /* this flag tells me if we are processing an event */

	
		while ( rowNumAt < values.size() ) {			  
			verboseMessage("PROCESSING ROW " + rowNumAt + "...");
			  
			List<Object> rowAt = values.get(rowNumAt);
			if ( rowAt.size() == 0 ) {
				/* Skip null rows */
				verboseMessage(" null row. ");
				  
				rowNumAt++;
			} else {
				/* Check and See if we are at a new event. */
				if ( inAnEventFlag == false ) {
                    verboseMessage("InAnEventFlag is false"); 
					if ( ! firstColOnlyContent(rowAt) ) {
						if ( garbageRecord(rowAt) ) {
							verboseMessage("Skipping a garbage record <" + rowNumAt + ">");
						} else {
							throw new ExcelExtractorException("Something weird PhysicalNumberOfCells is <" + rowAt.size() + "> expecting 0",rowNumAt);
						}	
					} else {
						String cellZero = rowAt.get(0).toString();
						if ( isMatchHeaderToken(rowAt) ) {
							throw new ExcelExtractorException("Something weird",rowNumAt);
						} else if ( cellZero.startsWith(OFFICIAL_TOKEN) ) {
							verboseMessage("Skip row, it is a lingering Official row from tourney.");
						} else {
						  /* 
						   * We are at a new event.  Next step is to see if it is a tournament or a dual.
						   */
						  String eventString = cellZero;
					      verboseMessage(" is a match <" + eventString + ">");
						  inAnEventFlag = true;
						   /*
						    * check to see if it is a tournament.
						    */
						    if ( isATourney(eventString) ) {
								Tournament t = initializeTourney(eventString);
								verboseMessage(" and a tourney confirmed. ");
								 /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								*/
							   int addRow = this.processTourneyHeaderRows(values);
							   rowNumAt += addRow; 
							   verboseMessage("----- Tourney:" + t);
							    /*
							    * Now we process the tourney itself.
							    */
							   addRow = this.processTourneyMatches(t,values);
							   rowNumAt += addRow;
							   verboseMessage("rowNumAt now " + rowNumAt);
							   if ( year == 1) {
								   theTeam.addPrestigeLastYearTournament(t);
							   } else if (year == 2 ) {
								   theTeam.addPrestige2YearsAgoTournament(t);
							   } else if ( year == 3 ) {
								   theTeam.addPrestige3YearsAgoTournament(t);
							   } else {
								   throw new ExcelExtractorException(" Prestige Year " + year + " is not acceptable!", rowNumAt);
							   }
							   inAnEventFlag=false;
						    }
						   /* 
						    * Something went wrong, throw an exception....
							*/
							else {
								throw new ExcelExtractorException("Can't determine if dual or tourney. ",rowNumAt);
						    }
						}
					}	
					rowNumAt++;
				} else {   /*here we are already in an event */
					if ( isMatchHeaderToken(rowAt) ){
						verboseMessage(" is a match token") ;
					} else {
						verboseMessage(" is not a match token");
				    }
					rowNumAt++;
				}
			}
		}
		syncPrestige(theTeam,year);
		return theTeam;
	}
	public Team extractTeam() throws Exception {

        Team theTeam = new Team();
        verboseMessage("Starting..."); // Display the string.
		verboseMessage("Working with file <" + this.getSheetId() + "> team <" + this.getTeam() + ">"); 
		  
		theTeam.setTeamName(team);

		/*
		 * Process last year roster first.
		*/
		theTeam = this.extractLastYearRoster(theTeam);
		  
		/*
		* then process the roster.
		*/
		theTeam = this.extractRoster(theTeam);

		/*
		 * Process Weigh In History.
		*/
		theTeam = extractWeighInHistory(theTeam);
	
		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
	
		theTeam = extractResults(theTeam);
		
		theTeam.syncWithLastYearRoster();
		theTeam.buildAllBoutsLookup();

		/*
		 * Process Last Year Prestige.
		*/
		theTeam = extractPrestige(theTeam,1);

		

		theTeam = extractPrestige(theTeam,2);
	
		theTeam = extractPrestige(theTeam,3);
		
		theTeam = extractLastYearResults(theTeam);
		

		return theTeam;
		
	}
	
    public Team extractLastYearResults(Team theTeam) throws Exception {
		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
		rowNumAt=0;
		
		String range = SHEET_LAST_YEAR_RESULTS + "!A:E";
		
		ValueRange response = service.spreadsheets().values()
	                .get(this.getSheetId(), range)
	                .execute();
	    List<List<Object>> values = response.getValues();
		
		if ( values == null || values.size() == 0 ) {
			return theTeam;
		}

		/*
		 * This while loop will process chunks of the results.  This is why the rowAt is not incremented
		 * in the main line of a for loop.
		 */
		boolean inAnEventFlag = false;   /* this flag tells me if we are processing an event */
		while ( rowNumAt < values.size() ) {			  
			verboseMessage("PROCESSING ROW ...");
			  
			List<Object> rowAt = values.get(rowNumAt);
			if ( rowAt.size() == 0 ) {
				/* Skip empty rows */
				verboseMessage(" empty row. ");
				rowNumAt++;
			} else {
				/* Check and See if we are at a new event. */
				if ( inAnEventFlag == false ) {
                    verboseMessage("InAnEventFlag is false"); 
					if ( ! firstColOnlyContent(rowAt) ) {
						if ( garbageRecord(rowAt) ) {
							verboseMessage("Skipping a garbage record");
						} else {
							throw new ExcelExtractorException("Something weird PhysicalNumberOfCells is <" + rowAt.size() + "> expecting 0",rowNumAt);
						}	
					} else {
						String cellZero = rowAt.get(0).toString();
						if ( isMatchHeaderToken(rowAt) ) {
							throw new ExcelExtractorException("Something weird",rowNumAt);
						} else if ( cellZero.startsWith(OFFICIAL_TOKEN) ) {
							verboseMessage("Skip row, it is a lingering Official row from dual.");
						} else if ( cellZero.startsWith(WINNING_TEAM_TOKEN) ) {
							verboseMessage("Skip row, winning team footer.");
						} else if ( cellZero.startsWith(CLICK_HERE_TOKEN) ) {
							verboseMessage("Skip row, click here row found.");
						} else {
						  /* 
						   * We are at a new event.  Next step is to see if it is a tournament or a dual.
						   */
						  String eventString = cellZero;
					      verboseMessage(" is a match <" + eventString + ">");
						  inAnEventFlag = true;
						  /*
						   *  Check to see if it is a dual
						   */
						   if ( this.isADual(eventString) ) {
							   verboseMessage(" and dual confirmed. ");
							   DualMeet d = this.initializeDual(eventString);
							   rowNumAt++;
							   /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								*/
							   this.processDualHeaderRows(values);
							   
							   verboseMessage("----- Dual:" + d );
							   /*
							    * Now we process the dual itself.
							    */
							   int addRow = this.processDualMatches(d,values);
							   rowNumAt += addRow;
							   verboseMessage("rowNumAt now " + rowNumAt);
							   /* look for trailer record. */
							   rowAt = values.get(rowNumAt);
							   if (rowAt != null) {
								   String c = rowAt.get(0).toString();
								   if ( c.equals(DUAL_FOOTER_USP)) {
									  verboseMessage("At Unsportsmanlike point");
									  rowNumAt++;
									  rowAt=values.get(rowNumAt);
									  c = rowAt.get(0).toString();

								   }
								   if ( c.equals(DUAL_FOOTER_TOKEN) ) {
									   verboseMessage("At Dual Footer row " );
								   } else {
									   rowNumAt--;
								   }
							   }
							   theTeam.addLYDualMeet(d);
							   inAnEventFlag=false;
						   } 
						   /*
						    * If it is not a dual, check to see if it is a tournament.
						    */
						    else if ( isATourney(eventString) ) {
								Tournament t = initializeTourney(eventString);
								verboseMessage(" and a tourney confirmed. ");
								 /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								*/
							   int addRow = this.processTourneyHeaderRows(values);
							   rowNumAt += addRow; 
							   verboseMessage("----- Tourney:" + t);
							    /*
							    * Now we process the tourney itself.
							    */
							   addRow = this.processTourneyMatches(t,values);
							   rowNumAt += addRow;
							   verboseMessage("rowNumAt now " + rowNumAt);
							   theTeam.addLYTournament(t);
							   inAnEventFlag=false;
						    }
						   /* 
						    * Something went wrong, throw an exception....
							*/
							else {
								throw new ExcelExtractorException("Can't determine if dual or tourney. ",rowNumAt);
						    }
						}
					}	
					rowNumAt++;
				} else {   /*here we are already in an event */
					if ( isMatchHeaderToken(rowAt) ){
						verboseMessage(" is a match token") ;
					} else {
						verboseMessage(" is not a match token");
				    }
					rowNumAt++;
				}
			}
		}
		return theTeam;
	}	
	/*
    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
 
       
        GoogleSheetsExtractor gex = new GoogleSheetsExtractor("Hunterdon Central Reg",spreadsheetId);
        try {
        	Team t = gex.extractTeam();
        	GoogleSheetsWriter g = new GoogleSheetsWriter("Hunterdon Central Reg",spreadsheetId);
        	g.scanSheets();
        	System.out.println("sheetid->" + g.getSheetId("GamePlan"));
        	g.writeTeam(t,t);
        } catch ( Exception e) { 
        	 e.printStackTrace(System.out);
        }
        
    }
    */
}