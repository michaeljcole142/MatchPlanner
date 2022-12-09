package statsProcessor;


import java.util.*;
import org.apache.poi.ss.usermodel.*;
import java.io.File;


/*  */



/**
 * This class is used to extract information that is in track format that
 * was copied into excel.  
 * 1) results
 * 2) roster
 * 3) last year results
 */
class ExcelExtractor {
 
 
	private static String DNP_TOKEN="DNP";
	private static String WEIGHIN_TOKEN="Weigh-ins";
	private static String DATE_TOKEN="Date";
	private static String MATCHES_TOKEN="Matches";
	private static String DASHBOARD_TOKEN="Dashboard";
	private static String TIE_BREAKING_TOKEN="Tie Breaking";
	private static String SUMMARY_TOKEN="Summary";
	private static String WEIGHTS_TOKEN="Weights";
	private static String DUAL_FOOTER_TOKEN="Team Score:";
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
  
	private String filename;
	private String team;
    private int rowNumAt;
	private boolean verbose=true;
	
    public ExcelExtractor(String t, String f) {
		setTeam(t);
		setFileName(f);
	}
	public String getFileName() { return filename; }
	public String getTeam() { return team; }
	
	public void setFileName(String f) { filename = f; }
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
	private GSDualMeet initializeDual(String eventString) throws ExcelExtractorException {
		String[] tokens = eventString.split(" vs. ");
		if (tokens.length != 2 ) {
			throw new ExcelExtractorException ("error in initializeDualString with " + eventString + " length is " + tokens.length,rowNumAt);
		}
		GSDualMeet theDual = new GSDualMeet();
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
	private boolean isMatchHeaderToken(Row thisRow ) {
		if ( thisRow.getFirstCellNum()== 0 ) {
			Cell firstCell = thisRow.getCell(0);
			String token = firstCell.getStringCellValue();
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
	private void processDualHeaderRows(Sheet resultSheet) throws ExcelExtractorException {
		Cell checkCell;
		boolean stillProcessing = true;
		while ( stillProcessing ) {
			Row nextRow = resultSheet.getRow(rowNumAt);
			if (nextRow == null) {
				verboseMessage("empty row");
				rowNumAt++;
			} else {
				checkCell = nextRow.getCell(0);
				if ( isABlankRow(nextRow)) {
					verboseMessage("skipping blank row");
					rowNumAt++;
				} else {
					if ( checkCell != null ) {
						if ( checkCell.getStringCellValue().equals(MATCHES_TOKEN) ||
							checkCell.getStringCellValue().equals(DASHBOARD_TOKEN) ||
							checkCell.getStringCellValue().equals(TIE_BREAKING_TOKEN) ||
							checkCell.getStringCellValue().equals(SUMMARY_TOKEN) ||
							checkCell.getStringCellValue().equals(WEIGHTS_TOKEN) ||
							checkCell.getStringCellValue().equals(DUALS_TOKEN) ||
							checkCell.getStringCellValue().equals(MORE_TOKEN) ) {
							verboseMessage("At Dual header record" ) ;
							rowNumAt++;
						}
					} else {
						Cell wCell = nextRow.getCell(1);
						if (wCell != null ) {
							if (wCell.getStringCellValue().equals(WEIGHT_TOKEN) ) {
								verboseMessage("Past Dual Headers");
								stillProcessing = false;
								rowNumAt--;
							}
						} else {
							throw new ExcelExtractorException("Unknown Dual Header <" + checkCell.getStringCellValue() + ">", rowNumAt);
					
						}
					}
				}
				
			}
		   
	   }
	 
	   /* Now should find dual header row. */
	   rowNumAt++;
	   Row dualHeaderRow = resultSheet.getRow(rowNumAt);
	   if ( dualHeaderRow == null ) {
		   throw new ExcelExtractorException("empty row when expected dual header row at " + rowNumAt,rowNumAt);
	   }
	   int lastCellNum = dualHeaderRow.getLastCellNum();
	   if ( lastCellNum != 5 ) {
		   throw new ExcelExtractorException ("Unexpected Header Count of " + lastCellNum + " at " + rowNumAt,rowNumAt);
	   }
	   
	   verboseMessage("Processed Dual Meet Headers Successfully.");
	   return;
  }	
	private int processTourneyHeaderRows(Sheet resultSheet, int rowAt) throws ExcelExtractorException {
       int rowCheck=1;
	   Cell checkCell;
	   Row teamsRow = resultSheet.getRow(rowAt+rowCheck);
	   if (teamsRow == null) {
		   throw new ExcelExtractorException("Didn't find Teams record. ",rowNumAt);
	   }
	   checkCell = teamsRow.getCell(0);
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Teams row was empty. ",rowNumAt);
	   }
	   if (! checkCell.getStringCellValue().equals(TEAMS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted non Teams record found.  " + checkCell.getStringCellValue(),rowNumAt );
       }
	   rowCheck++; rowCheck++;
       Row divisionsRow = resultSheet.getRow(rowAt+rowCheck);
	   if (divisionsRow == null) {
		   throw new ExcelExtractorException("Didn't find Divisions record. ",rowNumAt);
	   }
	   checkCell = divisionsRow.getCell(0);
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Divisions row was empty. ",rowNumAt);
	   }
	   if (!checkCell.getStringCellValue().equals(DIVISIONS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Tourney record found.  " + checkCell.getStringCellValue(),rowNumAt );
       }
	   rowCheck++;rowCheck++;
       Row weightsRow = resultSheet.getRow(rowAt+rowCheck);
	   if (weightsRow == null) {
		   throw new ExcelExtractorException("Didn't find Weights record. ",rowNumAt);
	   }
	   checkCell = weightsRow.getCell(0);
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Weights row was empty. ",rowNumAt);
	   }
	   if (!checkCell.getStringCellValue().equals(WEIGHTS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Weights record found.  " + checkCell.getStringCellValue(),rowNumAt );
       }
	   rowCheck++;rowCheck++;
       Row wrestlersRow = resultSheet.getRow(rowAt+rowCheck);
	   if (wrestlersRow == null) {
		   throw new ExcelExtractorException("Didn't find Wrestlers record. ",rowNumAt);
	   }
	   checkCell = wrestlersRow.getCell(0);
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("Wrestlers row was empty. ",rowNumAt);
	   }
	   if (!checkCell.getStringCellValue().equals(WRESTLERS_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Dual record found.  " + checkCell.getStringCellValue(),rowNumAt );
       }
	   rowCheck++;rowCheck++;
       Row moreRow = resultSheet.getRow(rowAt+rowCheck);
	   if (moreRow == null) {
		   throw new ExcelExtractorException("Didn't find More record. ",rowNumAt);
	   }
	   checkCell = moreRow.getCell(0);
	   if ( checkCell == null ) {
		   throw new ExcelExtractorException ("More row was empty. ",rowNumAt);
	   }
	   if (!checkCell.getStringCellValue().equals(MORE_TOKEN) ) {
		   throw new ExcelExtractorException( "Unexpedcted Dual record found.  " + checkCell.getStringCellValue(),rowNumAt );
       }
	   /* Should find a blank row next. */
	   rowCheck++;rowCheck++;
	   Row blankRow = resultSheet.getRow(rowAt+rowCheck);
	   if ( ! isABlankRow(blankRow) ) {
		   throw new ExcelExtractorException ("expecting blank row at " + (rowAt+rowCheck),rowNumAt);
	   }
	   /* Now should find tourney header row. */
	   rowCheck++;rowCheck++;
	   Row tourneyHeaderRow = resultSheet.getRow(rowAt+rowCheck);
	   if ( tourneyHeaderRow == null ) {
		   throw new ExcelExtractorException("empty row when expected tourney header row at " + (rowAt+rowCheck),rowNumAt);
	   }
	   int lastCellNum = tourneyHeaderRow.getLastCellNum();
	   if ( lastCellNum < 5 ) {
		   throw new ExcelExtractorException ("Unexpected Header Count of " + lastCellNum + " at " + (rowAt+rowCheck),rowNumAt);
	   }
	   
	   verboseMessage("Processed Tourney Headers Successfully."); 
	   return rowCheck-1;
  }	
	private int processDualMatches(GSDualMeet d, Sheet resultSheet, int rowAt) throws ExcelExtractorException {
       int rowCheck=1;
	   /*
	    *   Should find 14 weights in a dual.  
		*/
	   for ( int i=0; i<14; i++ ) {
			Row aRow = resultSheet.getRow(rowAt+rowCheck);
			Cell weightCell = aRow.getCell(1);
			Cell matchCell = aRow.getCell(2);
			Bout b = extractDualMatch(matchCell.getStringCellValue(),weightCell.getStringCellValue());
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
	private int processTourneyMatches(Tournament t, Sheet resultSheet, int rowAt) throws ExcelExtractorException {
       int rowCheck=1;
	   /*
	    *   Process until a blank row.  
		*/
		boolean atBlankRow = false;
		
	   while ( ! atBlankRow ) {
		   Row aRow = resultSheet.getRow(rowAt+rowCheck);
		   if ( aRow == null ) {
			   atBlankRow = true;
			   rowCheck--;
		   } else if ( aRow.getFirstCellNum() < 1 ) {
			   atBlankRow = true;
			   rowCheck--;
		   } else {
				Cell weightCell = aRow.getCell(1);
				Cell matchCell = aRow.getCell(2);
				
				System.out.println("---->"  + weightCell + " -->" + matchCell);
				Bout b = extractTourneyMatch(matchCell.getStringCellValue(),weightCell.getStringCellValue());
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
			System.out.println("ROWAT<" + rowNumAt + ">-" + message);
		}
	}
	private GSTeam extractRoster(Sheet rSheet,GSTeam t) throws ExcelExtractorException {
	    int firstRow = rSheet.getFirstRowNum();
		int lastRow = rSheet.getLastRowNum();
		verboseMessage("Roster: firstRow = " + firstRow + " last row = " + lastRow );
		/*
		 * This will read in an initialize the team roster.
		 * 
		 */
		rowNumAt=firstRow;
		while ( rowNumAt <=lastRow ) {			  
			verboseMessage("ROSTER: PROCESSING ROW " + rowNumAt + "...");
			Row rowAt = rSheet.getRow(rowNumAt);
			if ( rowAt ==  null ) {
				/* Skip null rows */
				verboseMessage("ROSTER: null row. ");  
			} else {
				verboseMessage("ROSTER: ROW<" + rowNumAt + "> Cells<" + rowAt.getLastCellNum() + ">");
				if ( rowAt.getLastCellNum() >= 1 ) {
					Cell nameCell = rowAt.getCell(3);
					if ( nameCell.getStringCellValue().equals(ROSTER_NAME_TOKEN) ) {
						verboseMessage("ROSTER: at header row.");
					} else {
					    Cell forceWtCell = rowAt.getCell(0);
					    Cell forceRankCell = rowAt.getCell(1);
						Cell wtClassCell = rowAt.getCell(5);
						Cell genderCell = rowAt.getCell(6);
						Cell gradeCell = rowAt.getCell(7);
						Cell recordCell = rowAt.getCell(8);
					
						String forceWtVal = "";
						Integer forceRankVal = null;
						String wtClassVal = "";
						String genderVal = "";
						String gradeVal = "";
						String recordVal = "";
					
						if ( forceWtCell != null ) { forceWtVal = forceWtCell.getStringCellValue(); }

						if ( forceRankCell != null ) { forceRankVal = Integer.parseInt(forceRankCell.getStringCellValue()); }
						if ( wtClassCell != null ) { wtClassVal = wtClassCell.getStringCellValue(); }
						if ( genderCell != null ) { genderVal = genderCell.getStringCellValue(); }
						if ( gradeCell != null ) { gradeVal = gradeCell.getStringCellValue(); }
						if ( recordCell != null ) { recordVal = recordCell.getStringCellValue(); }
					
						verboseMessage( nameCell.getStringCellValue() +
								wtClassVal + genderVal + gradeVal + recordVal );
				
						if ( t.wrestlerExists(nameCell.getStringCellValue()) ) {
							throw new ExcelExtractorException("ROSTER: wrestler " + nameCell.getStringCellValue() + " exists already.",rowNumAt);
						} else {
							GSWrestler aWrestler = new GSWrestler(nameCell.getStringCellValue(),team);
							aWrestler.setTrackWtClass(wtClassVal);
							aWrestler.setGender(getGender(genderVal));
							aWrestler.setGrade(getGrade(gradeVal));
							aWrestler.setTrackRecord(recordVal);
							aWrestler.setForceWeight(forceWtVal);
							aWrestler.setForceRank(forceRankVal);
							t.addWrestler(aWrestler);
						}
					}
				} else {
					verboseMessage("skipping row that has less than 1 cells.");
				}
			} 
			rowNumAt++;
		}
		return t;
	}
	private GSTeam extractWeighInHistory(Sheet rSheet,GSTeam t) throws ExcelExtractorException {
		if ( rSheet == null ) {
			return t;
		}
	    int firstRow = rSheet.getFirstRowNum();
		int lastRow = rSheet.getLastRowNum();
		verboseMessage("Weigh In History: firstRow = " + firstRow + " last row = " + lastRow );
		/*
		 * This will read in and Build a history then add it to a wrestler.
		 */
		
		rowNumAt=firstRow;
		GSWrestler wrestlerAt;
		while ( rowNumAt <= lastRow+1 ) {			  
			verboseMessage("WeighInHistory: PROCESSING ROW " + rowNumAt + " ...");
			Row rowAt = rSheet.getRow(rowNumAt);
			if ( rowAt ==  null ) {
				/* Skip null rows */
				verboseMessage("WeighInHistory: null row. ");  
			} else {
				verboseMessage("WeighInHistory: ROW<" + rowNumAt + "> Cells<" + rowAt.getLastCellNum() + ">");
				Cell firstCell = rowAt.getCell(0);
				if ( isABlankRow(rowAt) ) {
					/*not at a header row */
					verboseMessage("WeighInHistory: ROW<" + rowNumAt + "> skipping blank row. ");
				} else {
					String header = firstCell.getStringCellValue();
System.out.println("XXXXXXXXXXXXXXXXXX Wrestler " + header );
					String wrestlerName = header.substring(0,header.indexOf(WEIGHIN_TOKEN)-1);
System.out.println("XXXXXXXXXXXXXXXXXX Name is <" + wrestlerName + ">");
					wrestlerAt=t.getWrestler(wrestlerName);
					WeighInHistory wiHistory = new WeighInHistory();
					wiHistory.setName(wrestlerName);
					if ( wrestlerAt == null ) {
						throw new ExcelExtractorException("Can not find wrestler " + wrestlerName, rowNumAt );
					}
					verboseMessage("WeighInHistory: ROW<" + rowNumAt + "> <" + header + ">");
					rowNumAt++;
					boolean pHeaders = true;
					while ( pHeaders ) {
						rowAt=rSheet.getRow(rowNumAt);
						if ( rowAt==null ) {
							rowNumAt++;
						} else {
							firstCell = rowAt.getCell(0);
							Cell secondCell = rowAt.getCell(1);
							String firstVal="";
							String secondVal="";
							if ( firstCell != null ) { firstVal = firstCell.getStringCellValue(); }
							if ( secondCell != null ) { secondVal = secondCell.getStringCellValue(); }
							if ( firstVal != null && firstVal.length() > 0 ) {
								/* still at headers */
								rowNumAt++;
							} else if ( secondVal != null && secondVal.length() > 0 ) {
								/* now on a wi row. */
								pHeaders = false;
							} else if ( isABlankRow(rowAt) ) {
								verboseMessage("Skipping Blank Row at <" + rowNumAt + ">");
								rowNumAt++;
							} else {
								throw new ExcelExtractorException("WeighInHistory: At exception row. ", rowNumAt );
							}
						}
					}
					verboseMessage("WeighInHistory: Finished Header rows. ");
					/* processing weigh in records header now */
					Cell dtCell = rowAt.getCell(1);
					if ( dtCell == null ) { 
						throw new ExcelExtractorException("Unexpected empty cell at row <",rowNumAt);
					}
					String dt = dtCell.getStringCellValue();
					if ( ! dt.equals(DATE_TOKEN) ) {
						throw new ExcelExtractorException("Unexpected token <" + dt + ">", rowNumAt);
					}
System.out.println("processing " + dt );
					rowNumAt++;
					boolean pRow = true;
					while ( pRow ) {
						rowAt=rSheet.getRow(rowNumAt);
						if ( rowAt == null ) {
							/* skip empty row. */
							verboseMessage("skip null row at " + rowNumAt);
							rowNumAt++;
							wrestlerAt.setWeighInHistory(wiHistory);
							pRow=false;
						} else if (  isABlankRow(rowAt) ) {
							/* skip blank row. */
							verboseMessage("skip blank row at " + rowNumAt);
							rowNumAt++;
						} else {
							Cell ckCell = rowAt.getCell(0);
							if ( ckCell == null || ckCell.getStringCellValue().length() == 0 ) {
								WeighIn wi = new WeighIn();
								wi.setWIDate(rowAt.getCell(1).getStringCellValue());
								wi.setWIEvent(rowAt.getCell(2).getStringCellValue());
	
								String wiString="";
								Cell c3 = rowAt.getCell(3);
								if ( c3 != null ) {
									wiString = c3.getStringCellValue();
									if ( wiString.equals(DNP_TOKEN) ) {
										wi.setWIWeight(0);
									} else {
										Float wt = Float.valueOf(rowAt.getCell(3).getStringCellValue());
										wi.setWIWeight(wt);
									}
									Cell c4 = rowAt.getCell(4);
									if ( c4 != null ) {
										String ss = c4.getStringCellValue();
										if ( ss != null ) {
											wi.setEntryDateTime(ss);
										}
									}
									wiHistory.addWI(wi);
								}
								rowNumAt++;
							} else {
								System.out.println("ckCell = " + ckCell.getStringCellValue() );
								pRow = false;
								wrestlerAt.setWeighInHistory(wiHistory);
								rowNumAt--;
							}
						}
					
					}
				}
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
			return WrestlingLanguage.Grade.G5;
		} else if ( gs.equals ( GRADE_SO_TOKEN ) ) {
			return WrestlingLanguage.Grade.G6;
	    } else if ( gs.equals ( GRADE_JR_TOKEN ) ) {
			return WrestlingLanguage.Grade.G7;
		} else if ( gs.equals ( GRADE_SR_TOKEN ) ) {
			return WrestlingLanguage.Grade.G8; 
		} else {
			throw new ExcelExtractorException( "Unknown Grade Token found <" + gs + ">", rowNumAt);
		}
	}
	private GSTeam extractLastYearRoster(Sheet rSheet,GSTeam t) throws ExcelExtractorException {
	    int firstRow = rSheet.getFirstRowNum();
		int lastRow = rSheet.getLastRowNum();
		verboseMessage("Last Year Roster: firstRow = " + firstRow + " last row = " + lastRow );
		/*
		 * This will read in an initialize the team roster.
		 * 
		 */
		rowNumAt=firstRow;
		while ( rowNumAt <=lastRow ) {			  
			verboseMessage("LAST YEAR ROSTER: PROCESSING ROW " + rowNumAt + "..." );
			Row rowAt = rSheet.getRow(rowNumAt);
			if ( rowAt ==  null ) {
				/* Skip null rows */
				verboseMessage("LAST YEAR ROSTER: null row. ");  
			} else {
				verboseMessage("LAST YEAR ROSTER: ROW<" + rowNumAt + "> Cells<" + rowAt.getLastCellNum() + ">");
				Cell nameCell = rowAt.getCell(1);
				if ( nameCell.getStringCellValue().equals(ROSTER_NAME_TOKEN) ) {
					verboseMessage("LAST YEAR ROSTER: at header row.");
				} else {
					
					Cell wtClassCell = rowAt.getCell(3);
					Cell genderCell = rowAt.getCell(4);
					Cell gradeCell = rowAt.getCell(5);
					Cell recordCell = rowAt.getCell(6);
				
					String wtClassVal = "";
					String genderVal = "";
					String gradeVal = "";
					String recordVal = "";
					
					if ( wtClassCell != null ) { wtClassVal = wtClassCell.getStringCellValue(); }
					if ( genderCell != null ) { genderVal = genderCell.getStringCellValue(); }
					if ( gradeCell != null ) { gradeVal = gradeCell.getStringCellValue(); }
					if ( recordCell != null ) { recordVal = recordCell.getStringCellValue(); }
					
					verboseMessage( nameCell.getStringCellValue() +
								wtClassVal + genderVal + gradeVal + recordVal );
					if ( t.wrestlerLastYearExists(nameCell.getStringCellValue()) ) {
						throw new ExcelExtractorException("LAST YEAR ROSTER: wrestler " + nameCell.getStringCellValue() + " exists already.",rowNumAt);
					} else {
						GSWrestler aWrestler = new GSWrestler(nameCell.getStringCellValue(),team);
						aWrestler.printVerbose();
						aWrestler.setTrackWtClass(wtClassVal);
						aWrestler.setGender(getGender(genderVal));
						aWrestler.setGrade(getGrade(gradeVal));
						aWrestler.setTrackRecord(recordVal);
						t.addLastYearWrestler(aWrestler);
					}
				}
			} 
			rowNumAt++;
		}
		return t;
	}
	private boolean firstColOnlyContent(Row r) {
		int minColIdx = r.getFirstCellNum();
		int maxColIdx = r.getLastCellNum();
		for(int colIdx=minColIdx ; colIdx<maxColIdx ; colIdx++) {
			Cell cell = r.getCell(colIdx);
			if ( colIdx == 0) {
				if ( cell == null ) {
					return false;
				} else {
					if ( cell.getStringCellValue().length() <= 0 ) {
						return false;
					}
				}
			} else {
				if ( cell != null ) {
					if ( cell.getStringCellValue().length() > 0 ) {
						return false;
					}
				}
			}
		}
		return true;
	}
	/*
	 * This is used to skip rows at the footer of a dual that are probably
	 * points for misconduct. 
	 */
	private boolean garbageRecord(Row r) {

		Cell matchCell = r.getCell(2);
		if ( matchCell != null ) {
			if ( matchCell.getStringCellValue().length() == 0 ) {
				return true;
			}
		} else {
			return true;
		}
		return false;
	}	
		
	private boolean isABlankRow(Row r) {
		if ( r == null ) { return true; }
		int minColIdx = r.getFirstCellNum();
		int maxColIdx = r.getLastCellNum();
		
		for(int colIdx=minColIdx ; colIdx<maxColIdx ; colIdx++) {
			Cell cell = r.getCell(colIdx);
			if ( cell != null ) {
				if ( cell.getStringCellValue().length() > 0 ) {
					return false;
				}
			}
		}
		return true;
	}
    public GSTeam extractResults(GSTeam theTeam, Sheet resultsSheet) throws Exception {
		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
		int firstRow = resultsSheet.getFirstRowNum();
		int lastRow = resultsSheet.getLastRowNum();
		verboseMessage("firstRow = " + firstRow + " last row = " + lastRow );

		/*
		 * This while loop will process chunks of the results.  This is why the rowAt is not incremented
		 * in the main line of a for loop.
		 */
		boolean inAnEventFlag = false;   /* this flag tells me if we are processing an event */
		rowNumAt=firstRow;
		while ( rowNumAt <=lastRow ) {			  
			verboseMessage("PROCESSING ROW " + rowNumAt + "...");
			  
			Row rowAt = resultsSheet.getRow(rowNumAt);
			if ( rowAt ==  null ) {
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
							throw new ExcelExtractorException("Something weird PhysicalNumberOfCells is <" + rowAt.getPhysicalNumberOfCells() + "> expecting 0",rowNumAt);
						}	
					} else {
						Cell cellZero = rowAt.getCell(0);
						if ( isMatchHeaderToken(rowAt) ) {
							throw new ExcelExtractorException("Something weird",rowNumAt);
						} else if ( cellZero.getStringCellValue().startsWith(OFFICIAL_TOKEN) ) {
							verboseMessage("Skip row, it is a lingering Official row from dual.");
						} else {
						  /* 
						   * We are at a new event.  Next step is to see if it is a tournament or a dual.
						   */
						  String eventString = cellZero.getStringCellValue();
					      verboseMessage(" is a match <" + eventString + ">");
						  inAnEventFlag = true;
						  /*
						   *  Check to see if it is a dual
						   */
						   if ( isADual(eventString) ) {
							   verboseMessage(" and dual confirmed. ");
							   GSDualMeet d = initializeDual(eventString);
							   rowNumAt++;
							   /* This code makes sure the next 6 records are of dual format. 
							    * It will throw an exception if not.
								*/
							   processDualHeaderRows(resultsSheet);
							   
							   verboseMessage("----- Dual:" + d );
							   /*
							    * Now we process the dual itself.
							    */
							   int addRow = processDualMatches(d,resultsSheet,rowNumAt);
							   rowNumAt += addRow;
							   verboseMessage("rowNumAt now " + rowNumAt);
							   /* look for trailer record. */
							   rowAt = resultsSheet.getRow(rowNumAt);
							   if (rowAt != null) {
								   Cell c = rowAt.getCell(0);
								   if ( c.getStringCellValue().equals(DUAL_FOOTER_TOKEN) ) {
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
							   int addRow = processTourneyHeaderRows(resultsSheet,rowNumAt);
							   rowNumAt += addRow; 
							   verboseMessage("----- Tourney:" + t);
							    /*
							    * Now we process the tourney itself.
							    */
							   addRow = processTourneyMatches(t,resultsSheet,rowNumAt);
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
	private void setPrestige(GSWrestler w, String tourney, String round, WrestlingLanguage.WinOrLose worl, int year ) throws ExcelExtractorException {

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
	private void syncPrestige(GSTeam theTeam, int year) throws ExcelExtractorException {
	
		
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
				GSWrestler w = theTeam.getWrestler(name);
				if ( w != null ) {
					setPrestige(w,tourney,round,theBouts.get(ii).getWinOrLose(),year);
				}
			}
		}
		return;
	}
	public GSTeam extractPrestige(GSTeam theTeam, Sheet prestigeSheet,int year) throws Exception {
		/*
		 * This is the processing of last year's prestige sheet.
		 */		   
		int firstRow = prestigeSheet.getFirstRowNum();
		int lastRow = prestigeSheet.getLastRowNum();
		verboseMessage("firstRow = " + firstRow + " last row = " + lastRow );

		/*
		 * This while loop will process chunks of the results.  This is why the rowAt is not incremented
		 * in the main line of a for loop.
		 */
		boolean inAnEventFlag = false;   /* this flag tells me if we are processing an event */
		rowNumAt=firstRow;
		/* This if is a hack for bucks */
		if ( firstRow == lastRow ) { rowNumAt++;  }
		while ( rowNumAt <=lastRow ) {			  
			verboseMessage("PROCESSING ROW " + rowNumAt + "...");
			  
			Row rowAt = prestigeSheet.getRow(rowNumAt);
			if ( rowAt ==  null ) {
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
							throw new ExcelExtractorException("Something weird PhysicalNumberOfCells is <" + rowAt.getPhysicalNumberOfCells() + "> expecting 0",rowNumAt);
						}	
					} else {
						Cell cellZero = rowAt.getCell(0);
						if ( isMatchHeaderToken(rowAt) ) {
							throw new ExcelExtractorException("Something weird",rowNumAt);
						} else if ( cellZero.getStringCellValue().startsWith(OFFICIAL_TOKEN) ) {
							verboseMessage("Skip row, it is a lingering Official row from tourney.");
						} else {
						  /* 
						   * We are at a new event.  Next step is to see if it is a tournament or a dual.
						   */
						  String eventString = cellZero.getStringCellValue();
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
							   int addRow = processTourneyHeaderRows(prestigeSheet,rowNumAt);
							   rowNumAt += addRow; 
							   verboseMessage("----- Tourney:" + t);
							    /*
							    * Now we process the tourney itself.
							    */
							   addRow = processTourneyMatches(t,prestigeSheet,rowNumAt);
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
	public GSTeam extractTeam() throws Exception {

        GSTeam theTeam = new GSTeam();
        verboseMessage("Starting..."); // Display the string.
		verboseMessage("Working with file <" + filename + "> team <" + team + ">"); 
		Workbook workbook = WorkbookFactory.create(new File(filename));
		verboseMessage("Workbook has " + workbook.getNumberOfSheets());
        verboseMessage("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
			verboseMessage("=> " + sheet.getSheetName());
        }
		  
		theTeam.setTeamName(team);
		/*
		 * Process last year roster first.
		*/
		Sheet lastYearRosterSheet = workbook.getSheet("LastYearRoster");
		theTeam = extractLastYearRoster(lastYearRosterSheet,theTeam);
		rowNumAt = 0;  
		  
		/*
		* then process the roster.
		*/
		Sheet rosterSheet = workbook.getSheet("Roster");
		theTeam = extractRoster(rosterSheet,theTeam);
		rowNumAt = 0;  

		/*
		 * Process Weigh In History.
		*/
		Sheet weighInHistorySheet = workbook.getSheet("WeighInHistory");
		rowNumAt=0;  
		theTeam = extractWeighInHistory(weighInHistorySheet,theTeam);

		/*
		 * This is the processing of the sheet that has dual and tourney results for a team on it.
		 */		   
		Sheet resultsSheet = workbook.getSheet("Results");
		
		theTeam = extractResults(theTeam,resultsSheet);
		
		theTeam.syncWithLastYearRoster();
		theTeam.buildAllBoutsLookup();

		/*
		 * Process Last Year Prestige.
		*/
		Sheet prestigeLastYearSheet = workbook.getSheet("PrestigeLastYear");

		if ( prestigeLastYearSheet == null ) {
			System.out.println("No prestige last year sheet!");
		} else {
			rowNumAt=0;  
			theTeam = extractPrestige(theTeam,prestigeLastYearSheet,1);
		}
		Sheet prestige2YearsAgoSheet = workbook.getSheet("Prestige2YrsAgo");

		if ( prestige2YearsAgoSheet == null ) {
			System.out.println("No prestige 2 Years Ago sheet!");
		} else {
			rowNumAt=0;  
			theTeam = extractPrestige(theTeam,prestige2YearsAgoSheet,2);
		}
		Sheet prestige3YearsAgoSheet = workbook.getSheet("Prestige3YrsAgo");

		if ( prestige3YearsAgoSheet == null ) {
			System.out.println("No prestige 3 Years Ago sheet!");
		} else {
			rowNumAt=0;  
			theTeam = extractPrestige(theTeam,prestige3YearsAgoSheet,3);
		}
		return theTeam;
		
	}
}
