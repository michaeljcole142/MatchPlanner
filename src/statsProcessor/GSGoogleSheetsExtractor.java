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


public class GSGoogleSheetsExtractor {
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
        InputStream in = GSGoogleSheetsExtractor.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
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
	
	private static String ROSTER_NAME_TOKEN="Wrestler";
	private static String Result_NAME_TOKEN="Wrestler (H)";


	private static String GRADE_5_TOKEN="5";
	private static String GRADE_6_TOKEN="6";
	private static String GRADE_7_TOKEN="7";
	private static String GRADE_8_TOKEN="8";
	
	private static String GS_DUAL_HEADER_TOKEN="Varsity";
	private static String GS_WGT_TOKEN="WGT";
	
	// match results tokens
	private static String GS_DOUBLE_FORFEIT="Double Forfeit";
	private static String GS_WIN_DEC="Win (Decision)";
	private static String GS_WIN_MAJOR_DEC="Win (Major Decision)";
	private static String GS_WIN_TECH_FALL="Win (Technical Fall)";
	private static String GS_WIN_PIN="Win (Pin)";
	private static String GS_WIN_FFT="Win (Forfeit)";
	private static String GS_WIN_INJURY_DEFAULT="Win (Injury Default)";
	private static String GS_LOSS_DEC="Loss (Decision)";
	private static String GS_LOSS_MAJOR_DEC="Loss (Major Decision)";
	private static String GS_LOSS_TECH_FALL="Loss (Technical Fall)";
	private static String GS_LOSS_PIN="Loss (Pin)";
	private static String GS_LOSS_FFT="Loss (Forfeit)";
	private static String GS_LOSS_INJURY_DEFAULT="Loss (Injury Default)";

	private static String SHEET_RESULTS="Results";
	private static String SHEET_ROSTER="Roster";
		
	private String sheetId;
	private String team;
    private int rowNumAt;
	private boolean verbose=true;
	private Sheets service;

	
    public GSGoogleSheetsExtractor(String t, String sid) {
		setTeam(t);
		setSheetId(sid);
	    try {
	    	final NetHttpTransport HTTP_TRANSPORT = new com.google.api.client.http.javanet.NetHttpTransport();

	    	Sheets.Builder s2 = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT));
	    	s2.setApplicationName(APPLICATION_NAME);
	        service = s2.build();
	    } catch (Exception e) { 
	    	System.out.println("ERROR->" + e.toString());
	    }
	}
    
    public GSGoogleSheetsExtractor(String sid) {
		setSheetId(sid);
	    try {
	    	final NetHttpTransport HTTP_TRANSPORT = new com.google.api.client.http.javanet.NetHttpTransport();

	    	Sheets.Builder s2 = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT));
	    	s2.setApplicationName(APPLICATION_NAME);
	        service = s2.build();
	    } catch (Exception e) { 
	    	System.out.println("ERROR->" + e.toString());
	    }
	}    
	public String getSheetId() { return sheetId; }
	public String getTeam() { return team; }
	
	public void setSheetId(String sid) { sheetId = sid; }
	public void setTeam(String t) { team = t; }
	
	private boolean isADual(String eventString) {
		
		String[] tkns = eventString.split("\\|");
		if ( tkns.length != 3 ) {
			return false;
		}
		if ( tkns[1].split("vs.").length == 2 ) {
			return true;
		}
		return false;
	}
	private Bout extractDualMatch(GSDualMeet d, String weight,String wrestlerH, String wrestlerV, String resultCell, String boutScore, String boutTime, String matchName) throws ExcelExtractorException {
		Bout theBout = new Bout();
		theBout.setWeight(weight);
	
		String homeTeam ="", visitorTeam="";
		
		String wrestler1 = wrestlerH;
        wrestler1=wrestler1.trim();
        String wrestler2 = wrestlerV;
        wrestler2=wrestler2.trim();
        
        if (matchName.contains("Varsity Match")) {
			String VS = matchName.split("\\|")[1];
			 homeTeam = VS.split("vs.")[0].trim();
			 visitorTeam = VS.split("vs.")[1].trim();
			 homeTeam=homeTeam.trim();		
        } else {
			throw new ExcelExtractorException( "Unknown match name encountered... ",rowNumAt );
        }

		theBout.setResult(resultCell);
			
		boolean isAWin;
		
		if ( resultCell.equals(GS_DOUBLE_FORFEIT) ) {
			return null;
		} else if ( resultCell.equals(GS_WIN_DEC) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.DEC);
			theBout.setResult( WrestlingLanguage.MatchResultType.DEC + "(" + boutScore + ")");
			isAWin=true;
		} else if ( resultCell.equals(GS_WIN_MAJOR_DEC) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.MD) ;
			theBout.setResult( WrestlingLanguage.MatchResultType.MD + "(" + boutScore + ")");
			isAWin=true;
		} else if ( resultCell.equals(GS_WIN_TECH_FALL) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.TECH);
			theBout.setResult( WrestlingLanguage.MatchResultType.TECH + "(" + boutScore + " " + boutTime +  ")");
			isAWin=true;
		} else if ( resultCell.equals(GS_WIN_PIN) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FALL);
			theBout.setResult( WrestlingLanguage.MatchResultType.FALL + "(" + boutTime +  ")");
			isAWin=true;
		} else if ( resultCell.equals(GS_WIN_FFT) ) {
			if ( ! d.getIsHome() ) { return null; }
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FFT);
			theBout.setResult(WrestlingLanguage.MatchResultType.FFT + "");
			isAWin=true;
		} else if ( resultCell.equals(GS_WIN_INJURY_DEFAULT) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.INJ);
			theBout.setResult( WrestlingLanguage.MatchResultType.INJ + "(" +  boutTime +  ")");
			isAWin=true;
		} else if ( resultCell.equals(GS_LOSS_DEC) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.DEC);
			theBout.setResult( WrestlingLanguage.MatchResultType.DEC + "(" + boutScore + ")");
			isAWin=false;
		} else if ( resultCell.equals(GS_LOSS_MAJOR_DEC) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.MD) ;
			theBout.setResult( WrestlingLanguage.MatchResultType.MD + "(" + boutScore + ")");
			isAWin=false;
		} else if ( resultCell.equals(GS_LOSS_TECH_FALL) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.TECH);
			theBout.setResult( WrestlingLanguage.MatchResultType.TECH + "(" + boutScore + " " + boutTime +  ")");
			isAWin=false;
		} else if ( resultCell.equals(GS_LOSS_PIN) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FALL);
			theBout.setResult( WrestlingLanguage.MatchResultType.FALL + "(" + boutTime +  ")");
			isAWin=false;
		} else if ( resultCell.equals(GS_LOSS_FFT) ) {
			if ( d.getIsHome() ) { return null; } 
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.FFT);
			theBout.setResult(WrestlingLanguage.MatchResultType.FFT + "");
			isAWin=false;
		} else if ( resultCell.equals(GS_LOSS_INJURY_DEFAULT) ) {
			theBout.setMatchResultType(WrestlingLanguage.MatchResultType.INJ);
			theBout.setResult( WrestlingLanguage.MatchResultType.INJ + "(" +  boutTime +  ")");
			isAWin=false;
		} else {
			throw new ExcelExtractorException( "Unknown match result type encountered... <" + resultCell + ">",rowNumAt );
		}
		if ( ! d.getIsHome() ) {
			isAWin = ! isAWin;
		}
		if ( isAWin ) {
			theBout.setWin();
		} else {
			theBout.setLoss();
		}
		if ( d.getIsHome() ) {
			theBout.setMainName(wrestler1);
			theBout.setMainTeam(homeTeam);
			theBout.setOpponentName(wrestler2);
			theBout.setOpponentTeam(visitorTeam);
		} else {
			theBout.setMainName(wrestler2);
			theBout.setMainTeam(visitorTeam);
			theBout.setOpponentName(wrestler1);
			theBout.setOpponentTeam(homeTeam);
		}  
		
		return theBout;
	}
	
	private GSDualMeet initializeDual(String eventString) throws ExcelExtractorException {
		String[] tokens = eventString.split(" vs. ");
		if (tokens.length != 2 ) {
			throw new ExcelExtractorException ("error in initializeDualString with " + eventString + " length is " + tokens.length,rowNumAt);
		}
		GSDualMeet theDual = new GSDualMeet();
		theDual.setMainTeam(team);
		String VS = eventString.split("\\|")[1];
		String dt = eventString.substring(14,24);
		String homeTeam = VS.split("vs.")[0].trim();
		String visitorTeam = VS.split("vs.")[1].trim();
		if ( homeTeam.equals(team) ) { 
			theDual.setOpponent(visitorTeam);
			theDual.setIsHome();
		} else {
			theDual.setOpponent(homeTeam);
			theDual.setIsAway();
		}
		theDual.setEventDate(dt);
		theDual.setEventTitle(visitorTeam + " Dual");
		
		return theDual;
	}

	
	private boolean isMatchHeaderToken(List<Object> thisRow ) {
		if ( thisRow.size() >= 1 ) {
			String token = thisRow.get(0).toString().substring(0,GS_DUAL_HEADER_TOKEN.length());
			if ( token.equals(GS_DUAL_HEADER_TOKEN)  ) {
				verboseMessage("Token is match header " + token); 
				return true;
			}
		}
		verboseMessage("Token is not a match header "); 
		return false;
	}
	private void processDualHeaderRows(List<List<Object>> resultSheet) throws Exception {
		System.out.println("processing->" + rowNumAt);
			
		List<Object> nextRow = resultSheet.get(rowNumAt);
						   	 
	   if ( nextRow == null ) {
		   throw new ExcelExtractorException("empty row when expected dual header row at " + rowNumAt,rowNumAt);
	   }
	   int lastCellNum = nextRow.size()-1;
	   if ( lastCellNum != 7 ) {
		   throw new ExcelExtractorException ("Unexpected Header Count of " + lastCellNum + " at " + rowNumAt,rowNumAt);
	   }
	   if ( ! nextRow.get(0).toString().equals(GS_WGT_TOKEN)) {
		   throw new ExcelExtractorException("Unexpected Headre Token at" + rowNumAt,rowNumAt);
	   }
	   rowNumAt++;
	   verboseMessage("Processed Dual Meet Headers Successfully.");
	   return;
  }	
	
	private int processDualMatches(GSDualMeet d, List<List<Object>> resultSheet, String eventString) throws ExcelExtractorException {
       int rowCheck=0;
	
       System.out.println("size->" + d.getDualSize());
	   for ( int i=0; i< d.getDualSize(); i++ ) {
		   System.out.println("getting->" + rowNumAt + " ->"  + rowCheck);

			List<Object> aRow = resultSheet.get(rowNumAt+rowCheck);
			String weightCell = aRow.get(0).toString();
			String wrestlerH = aRow.get(1).toString();
			String wrestlerV = aRow.get(2).toString();
			String result = aRow.get(3).toString();
			String boutScore = aRow.get(4).toString();
			String boutTime="";
			if ( aRow.size() >= 8 ) {
				boutTime = aRow.get(7).toString();
			} 
			System.out.println("Team->" + d.getMainTeam() + " opp->" + d.getOpponent());

			Bout b = extractDualMatch(d,weightCell,wrestlerH, wrestlerV, result,boutScore,boutTime,eventString);
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
	
    private void verboseMessage(String message) {
		if ( verbose ) {
			System.out.println("ROWAT<" + (rowNumAt+1) + ">-" + message);
		}
	}
	private GSTeam extractRoster(GSTeam t) throws Exception {

		rowNumAt=0;
		//Setting the range value
		String rangeRoster = SHEET_ROSTER + "!A:E";
		
		ValueRange responseRoster = service.spreadsheets().values()
	                .get(this.getSheetId(), rangeRoster)
	                .execute();
	    List<List<Object>> valuesRoster = responseRoster.getValues();
	
	    //Printing the number of cells with value
	    verboseMessage("Roster: " + valuesRoster.size());;
		/*
		 * This will read in an initialize the team roster.
		 * Loops through all the wrestler to create a team
		 */
		while ( rowNumAt < valuesRoster.size() ) {			  
			verboseMessage("ROSTER: PROCESSING ROW " + rowNumAt + "..." );
			List<Object> rowAt = valuesRoster.get(rowNumAt);
			if ( rowAt ==  null ) {
				/* Skip null rows */
				verboseMessage("ROSTER: null row. ");  
			} else {
				verboseMessage("ROSTER: ROW<" + rowNumAt + "> Cells<" + rowAt.size() + ">");
				if ( rowAt.size() >= 5 ) {
					String nameCell = rowAt.get(1).toString();
					if ( nameCell.equals(ROSTER_NAME_TOKEN) ) {
						verboseMessage("ROSTER: at header row.");
					} else {
						//Creating strings to hold the values
						
						String jvOrVarsity = rowAt.get(0).toString();
						String nameval = rowAt.get(1).toString();
						String gradeVal = rowAt.get(2).toString();
						String seedval = rowAt.get(3).toString();
						String certval = rowAt.get(4).toString();
						verboseMessage( nameval + gradeVal + seedval + certval);
						if ( t.wrestlerExists(nameCell) ) {
							throw new ExcelExtractorException("ROSTER: wrestler " + nameCell + " exists already.",rowNumAt);
						} else {
							//putting values into wrestler objects.
							GSWrestler aWrestler = new GSWrestler(nameCell,team);
							aWrestler.setGrade(getGrade(gradeVal));
							aWrestler.setSeed(seedval);
							aWrestler.setCert(certval);
							aWrestler.setJVOrVarsity(jvOrVarsity);
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
	
	public GSTeam extractResultsForDual(GSTeam theTeam) throws Exception {
		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
		rowNumAt=0;
		
		String range = SHEET_RESULTS + "!A:H";
		
		ValueRange response = service.spreadsheets().values()
	                .get(this.getSheetId(), range)
	                .execute();
	    List<List<Object>> values = response.getValues();
		
		if ( values.size() == 0 ) {
			return theTeam;
		}
		String eventString="";

		/*
		 * This while loop will process chunks of the results.  This is why the rowAt is not incremented
		 * in the main line of a for loop.
		 */
		while ( rowNumAt < values.size() ) {	
			
			verboseMessage("PROCESSING ROW ...");
			  
			List<Object> rowAt = values.get(rowNumAt);
			if ( rowAt.size() == 0 ) {
				/* Skip empty rows */
				verboseMessage(" empty row. ");
				
			} else {
				
				if (rowAt.get(0).toString().contains("Varsity Match")) 
				{
					verboseMessage("Varsity Match Row");	
					continue;
				} else if ( rowAt.size() >= 7 ) {		
					
					String headerCell = rowAt.get(1).toString();

					if ( headerCell.equals(Result_NAME_TOKEN) ) {
						verboseMessage("RESLUT: at header row.");
					} else {
				
					      verboseMessage(" is a match <" + eventString + ">");
						  /*
						   *  Check to see if it is a dual
						   */
						   if ( this.isADual(eventString) ) {
							   verboseMessage(" and dual confirmed. ");
							   GSDualMeet d = this.initializeDual(eventString);
							   /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								
								*/							   
							   verboseMessage("----- Dual:" + d );
							   /*
							    * Now we process the dual itself.
							    */
							   int addRow = this.processDualMatches(d, values, eventString);
							   rowNumAt += addRow;
							   
							   verboseMessage("rowNumAt now " + rowNumAt);

							   theTeam.addDualMeet(d);
						   } 
					}
				}
				/* Check and See if we are at a new event. */
				
				
		}
		
		rowNumAt++;
		}
		return theTeam;
	}


    private WrestlingLanguage.Grade getGrade(String gs ) throws ExcelExtractorException {
		if ( gs.equals ( GRADE_5_TOKEN ) ) {
			return WrestlingLanguage.Grade.G5;
		} else if ( gs.equals ( GRADE_6_TOKEN ) ) {
			return WrestlingLanguage.Grade.G6;
	    } else if ( gs.equals ( GRADE_7_TOKEN ) ) {
			return WrestlingLanguage.Grade.G7;
		} else if ( gs.equals ( GRADE_8_TOKEN ) ) {
			return WrestlingLanguage.Grade.G8; 
		} else {
			throw new ExcelExtractorException( "Unknown Grade Token found <" + gs + ">", rowNumAt);
		}
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
		
	public GSTeam extractTeam() throws Exception {

        GSTeam theTeam = new GSTeam();
       // Team theTeam2 = new Team();
        verboseMessage("Starting..."); // Display the string.
		verboseMessage("Working with file <" + this.getSheetId() + "> team <" + this.getTeam() + ">"); 
		  
		theTeam.setTeamName(team);

		/*
		 * Process last year roster first.
		*/
		//theTeam = this.extractLastYearRoster(theTeam);
		  
		/*
		* then process the roster.
		*/
		theTeam = this.extractRoster(theTeam);
		
		//theTeam = this.extractAllTeams();


		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
	
		theTeam = this.extractResults(theTeam);
		
		theTeam.syncWithLastYearRoster();
		theTeam.buildAllBoutsLookup();


		return theTeam;
		
	}

    public GSTeam extractResults(GSTeam theTeam) throws Exception {
		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
		rowNumAt=0;
		
		String range = SHEET_RESULTS + "!A:H";
		
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
						if ( ! isMatchHeaderToken(rowAt) ) {
							verboseMessage("Skip row, it is not a dual.");
						} else {
						  /* 
						   * We are at a new event. We only have duals for GS.
						   */
						  String eventString = cellZero;
					      verboseMessage(" is a match <" + eventString + ">");
						  inAnEventFlag = true;
						  /*
						   *  Check to see if it is a dual
						   */
						   if ( this.isADual(eventString) ) {
							   verboseMessage(" and dual confirmed. ");
							   GSDualMeet d = this.initializeDual(eventString);
							   rowNumAt++;
							   /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								*/
							   this.processDualHeaderRows(values);
							   
							   verboseMessage("----- Dual:" + d );
							   /*
							    * Now we process the dual itself.
							    */
							   int addRow = this.processDualMatches(d,values,eventString);
							   rowNumAt += addRow;
							   verboseMessage("rowNumAt now " + rowNumAt);							  
							   theTeam.addDualMeet(d);
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
}