package statsProcessor;


import java.util.*;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

/**
 * This class is used to write to googlesheets any of the data that has been
 * accumulated.
 */
class GSGoogleSheetsWriter  {
 
  
	private String sheetId;
	private String team;
	private boolean verbose=true;
	private boolean overwrite=true;
	private Sheets service;

	public static String GAMEPLAN_SHEET="GamePlan";
	public static String VERBOSE_TEAM_SHEET="VerboseTeam";
	private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
	private static final String CREDENTIALS_FILE_PATH = "./credentials.json";
    private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

	private Color greyColor = new Color();
	private Color blueColor = new Color();
	private Color redColor = new Color(); 

	private static final String GS_JVV_VARSITY="Varsity";
	private static final String GS_JVV_JV="JV";
	private static final String GS_JVV_EXHIBITION="Exhibition";
	private static final String GS_JVV_NOSEED="NoSeed";

	
    public GSGoogleSheetsWriter(String t, String sid) {
		setTeam(t);
		setSheetId(sid);
		try {
		    final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		   this.service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
		               .setApplicationName(APPLICATION_NAME)
		               .build();
		   
		   Spreadsheet sp = service.spreadsheets().get(this.getSheetId()).execute();
		   List<Sheet> sheets = sp.getSheets();
		   System.out.println("XXXXXXXXXXXXXXXXXx->" + sheets.get(0).getProperties().getTitle());
		   //setting the colors
		   greyColor.setRed(Float.valueOf("50")); greyColor.setBlue(Float.valueOf("50"));greyColor.setGreen(Float.valueOf("50"));
			blueColor.setRed(Float.valueOf("0.0")); blueColor.setBlue(Float.valueOf("1.0"));blueColor.setGreen(Float.valueOf("0.0"));
			redColor.setRed(Float.valueOf("1.0")); redColor.setBlue(Float.valueOf("0.0"));redColor.setGreen(Float.valueOf("0.0"));
	   } catch (Exception e) { 
		   	System.out.println("ERROR->" + e.toString());
	   }
	}
    public void scanSheets() {
    	try {
		 
		   Spreadsheet sp = this.service.spreadsheets().get(this.getSheetId()).execute();
		   List<Sheet> sheets = sp.getSheets();
		   for ( int i=0; i < sheets.size(); i++ ) {
			   System.out.println("Sheet->" + sheets.get(i).getProperties().getTitle() + "  id->" + sheets.get(i).getProperties().getSheetId() );
	  
		   }
		} catch (Exception e) { 
		   	System.out.println("ERROR->" + e.toString());
	   }
    	
    }
    public Integer getSheetId(String name) {
    	try {
		 
		   Spreadsheet sp = this.service.spreadsheets().get(this.getSheetId()).execute();
		   List<Sheet> sheets = sp.getSheets();
		   for ( int i=0; i < sheets.size(); i++ ) {
			   String at = sheets.get(i).getProperties().getTitle();
			   if ( at.contentEquals(name)) {
				   Integer ret= sheets.get(i).getProperties().getSheetId();
				   return ret;
			   }
	  
		   }
		} catch (Exception e) { 
		   	System.out.println("ERROR->" + e.toString());
	   }
       return Integer.valueOf(-1);
    }
	public String getSheetId() { return sheetId; }
	public String getTeam() { return team; }
	
	public void setSheetId(String sid) { sheetId = sid; }
	public void setTeam(String t) { team = t; }
	
	public boolean isOverwriteOn() {
		return overwrite;
	}
	
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = GSGoogleSheetsWriter.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
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
	
	/* */
	private void verboseMessage(String m ) {
		if ( verbose ) {
			System.out.println(m);
		}
	}

	public void writeRoster(GSTeam oppTeam,GSTeam homeTeam) throws Exception {
	
		/* If GamePlan exists, delete it and start fresh. */
		if ( this.getSheetId(GAMEPLAN_SHEET) > 0 )  {
			verboseMessage("Gameplan Sheet found at index " + this.getSheetId(GAMEPLAN_SHEET) );
			List<Request> delRequests = new ArrayList<Request>();
			DeleteSheetRequest delSheet = new DeleteSheetRequest();
			delSheet.setSheetId(this.getSheetId(GAMEPLAN_SHEET));
			    delRequests.add(new Request().setDeleteSheet(delSheet));
				BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
				requestBody.setRequests(delRequests);
				service.spreadsheets().batchUpdate(this.getSheetId(), requestBody).execute();
				verboseMessage("Gameplan Sheet deleted." );
		}
		GridCoordinate startGrid = new GridCoordinate().setSheetId(this.getSheetId(GAMEPLAN_SHEET)).setRowIndex(0).setColumnIndex(0);
		UpdateCellsRequest body = new UpdateCellsRequest().setStart(startGrid).setFields("*");
		List<RowData> rows = new ArrayList<RowData>();

		Hashtable<String,GSWrestler> theRoster = oppTeam.getRoster();					
	   	Collection<GSWrestler> wc = theRoster.values();
	   	ArrayList<GSWrestler> aw = new ArrayList<GSWrestler> (wc);	
	   	Comparator<GSWrestler> compareIt = new Comparator<GSWrestler>() {
				public int compare(GSWrestler w1, GSWrestler w2) {
					
					int wS= w1.getFinWeight();
					int w2S= w2.getFinWeight();
					
					
					int i = wS - w2S;
				
					if ( i == 0 ) {
						i = (w2.getWins() + w2.getLosses()) - (w1.getWins() + w1.getLosses());
						if ( i == 0 ) {
							int w1jvv=0;
							int w2jvv=0;
							if ( w1.getJVOrVarsity().equals(GS_JVV_VARSITY)) {
								w1jvv=4;
							} else if ( w1.getJVOrVarsity().contentEquals(GS_JVV_JV)) {
								w1jvv=3;
							} else if ( w1.getJVOrVarsity().contentEquals(GS_JVV_EXHIBITION)) {
								w1jvv=2;
							} else if ( w1.getJVOrVarsity().contentEquals(GS_JVV_NOSEED)) {
								w1jvv=1;
							} 
							System.out.println("working on w2->" + w2);
							if ( w2.getJVOrVarsity().equals(GS_JVV_VARSITY)) {
								w2jvv=4;
							} else if ( w2.getJVOrVarsity().contentEquals(GS_JVV_JV)) {
								w2jvv=3;
							} else if ( w2.getJVOrVarsity().contentEquals(GS_JVV_EXHIBITION)) {
								w2jvv=2;
							} else if ( w2.getJVOrVarsity().contentEquals(GS_JVV_NOSEED)) {
								w2jvv=1;
							} 
							i = w2jvv - w1jvv;
							if ( i == 0 ) {
								if ( ! w2.getGradeString().equals("") && ! w1.getGradeString().equals("") ) {
									i = Integer.parseInt(w2.getGradeString()) - Integer.parseInt(w1.getGradeString());
								}
							}
						}	
					}
					return i;
				}
			};
			Collections.sort(aw,compareIt);
			
		List<CellData> r1 = new ArrayList<CellData>();
		r1.add(new CellData());
		CellData titleCell =new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(homeTeam.getTeamName() + " vs. " + oppTeam.getTeamName() + " Overview"));
		r1.add(titleCell);
	    CellFormat titleFormat = new CellFormat();
	    titleCell.setUserEnteredFormat(titleFormat);
	    TextFormat f = new TextFormat();
	    f.setFontSize(25);
	    f.setBold(true);
	    titleFormat.setTextFormat(f);
	
		rows.add(new RowData().setValues(r1));

		List<CellData> blankrow = new ArrayList<CellData>();
		rows.add(new RowData().setValues(blankrow));
		rows.add(new RowData().setValues(blankrow));

		List<CellData> headerrow = new ArrayList<CellData>();
		rows.add(new RowData().setValues(headerrow));
			
		CellFormat headerFormat = new CellFormat();
		TextFormat hf = new TextFormat();
		hf.setBold(true);
		headerFormat.setTextFormat(hf);		
		headerFormat.setWrapStrategy("WRAP");

		Borders hBorders = new Borders();
		hBorders.setTop(new Border().setStyle("SOLID_THICK"));
		hBorders.setBottom(new Border().setStyle("SOLID_THICK"));
		headerFormat.setBorders(hBorders);
			
		//creating header record
		CellData weightH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Weight"));
		headerrow.add(weightH);
		CellFormat leftHeaderFormat = headerFormat.clone();
		leftHeaderFormat.getBorders().setLeft(new Border().setStyle("SOLID_THICK"));
		weightH.setUserEnteredFormat(leftHeaderFormat);
	
		CellData nameH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Name"));
		headerrow.add(nameH);
		nameH.setUserEnteredFormat(headerFormat);
		CellData gradeH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Grade"));
		headerrow.add(gradeH);
		gradeH.setUserEnteredFormat(headerFormat);
		
		CellData jvvH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("V/J"));
		headerrow.add(jvvH);
		jvvH.setUserEnteredFormat(headerFormat);
		
		
		CellData recordH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Record"));
		headerrow.add(recordH);
		recordH.setUserEnteredFormat(headerFormat);
		CellData breakdownH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("BreakDown"));
		headerrow.add(breakdownH);
		breakdownH.setUserEnteredFormat(headerFormat);
		CellData lastYearH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Last Year"));
		headerrow.add(lastYearH);
		lastYearH.setUserEnteredFormat(headerFormat);
		CellData notesH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Notes"));
		headerrow.add(notesH);
		notesH.setUserEnteredFormat(headerFormat);
		CellData initialWIH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Cert Weight"));
		headerrow.add(initialWIH);
		initialWIH.setUserEnteredFormat(headerFormat);
		CellData seedWt = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Seed Weight"));
		headerrow.add(seedWt);
		
		CellFormat rightHeaderFormat = headerFormat.clone();
		rightHeaderFormat.getBorders().setRight(new Border().setStyle("SOLID_THICK"));
		seedWt.setUserEnteredFormat(rightHeaderFormat);

		//end of header record
		
		int count=0;
		String atWeightStr="";
		boolean oddWeight=true;
		
		
		boolean atNewWeight=true;
		int[] allWeights=GSDualMeet.getWeightListInt();
		for ( int ii=0; ii < allWeights.length; ii++ ) {

		atNewWeight=true;
		boolean found = false;
		if ( ii%2 ==1 ) {
			oddWeight=true;
		} else {
			oddWeight=false;
		}
		while ( aw.size() > count ) {
			GSWrestler w = aw.get(count);
			if ( w.getFinWeight() == allWeights[ii]) {

				System.out.println("w=" + w);	
				String wWeight = String.valueOf(w.getFinWeight());
			
				String weightData="";
				String nameData="";
				String gradeData="";
				String vjvData="";
				String recordData="";
				String breakdownData="";
				String lastYearData="";
				String notesData="";
				String certWeightData="";
				String seedWeightData="";
				Boolean defaultFont=false;
			
				/* Check and see if there is a wrestler for the weight 
				 * If so, set the data to be the current wrestler. 
				 * If not, set a blank row for that weight.
				 */
				weightData=String.valueOf(w.getFinWeight()) ;
				nameData=w.getName();

				gradeData=w.getGradeString();
				vjvData=w.getJVOrVarsity();
				recordData=w.getRecordString();
				
				breakdownData="";
				if ( w.getRecordBreakdown().length() > 0 ) {
					breakdownData = w.getRecordBreakdown() + ";" + w.getMatchesAtWeightString();
				} 
				
				lastYearData=w.getTrackLastYearRecord();
				
				List<Bout> bouts = w.getBouts();
				String h2h="";
				String inj="";
				String common="";
				int cSize=0;

				if ( bouts != null ) {
					for ( int i=0; i < bouts.size(); i++ ) {
						Bout b = bouts.get(i);
		
						if ( b.getOpponentTeam().equals(homeTeam.getTeamName())) {
							if ( h2h.length() > 0 ) { h2h += "\n"; }
						
							h2h += b.getWinOrLose() + " " + b.getResult() + " " +  b.getOpponentName();
						}
			
						String k = b.getOpponentTeam() + ":" + b.getOpponentName();
			
						List<Bout> bb = homeTeam.lookupCommonMatch(k);
			
						if ( bb != null ) {
							if  ( bb.size() > 0 ) {
								cSize += bb.size();
							}
						}
					}
					if ( cSize > 0 ) { 
						if ( common.length() > 0 ) { common += "\n"; }
						common += cSize + " common opponents";
					}	
				}
				notesData="";
				if ( h2h.length() > 0 ) {
					if (common.length() > 0 ) {
						notesData = h2h + "\n" + common;
					} else {
						notesData = h2h;
					}
				} else {
					if ( common.length() > 0 ) {
						notesData = common;
					}
				}
				if ( w.getLossByInjury() ) {
					defaultFont=true;
					if ( notesData.length() > 0 ) {
						notesData = notesData + "\n" + w.getLossByInjuryString();
					} else {
						notesData = w.getLossByInjuryString();
					}
				}
				certWeightData=w.getCert();
				seedWeightData=w.getSeed();

				
				List<CellData> row = writeRosterRow(atNewWeight, oddWeight, defaultFont, weightData, 
					 nameData,  gradeData,  vjvData,  recordData, 
					 breakdownData, lastYearData, notesData,  certWeightData,  seedWeightData);
			
				rows.add(new RowData().setValues(row));
				atNewWeight=false;
				count ++;
				found=true;
			} else {
				break;
			}
		}
		if ( ! found ) {
			if ( allWeights[ii] != 9999 ) {
				List<CellData> row = writeRosterRow(atNewWeight, oddWeight, false, String.valueOf(allWeights[ii]), "",  "",  "",  "",  "", "", "",  "",  "");
				rows.add(new RowData().setValues(row));
			}
		}
	}
	if ( count < aw.size()-1 ) {
		System.out.println("leftovers->" + count + " for->" + aw.size());
		
	}
		//Creating the new sheet
		AddSheetRequest addSheet = new AddSheetRequest();
		addSheet.setProperties(new SheetProperties().setTitle(GAMEPLAN_SHEET));
		List<Request> requests = new ArrayList<Request>();
		    
		requests.add(new Request().setAddSheet(addSheet));
		BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
		requestBody.setRequests(requests);

		service.spreadsheets().batchUpdate(this.getSheetId(), requestBody).execute();
				
		/* Setting the body data of the gameplan sheet. */
		body.setRows(rows);
		List<Request> updateCellsRequestList = new ArrayList<Request>();

		/* We created a new gameplan sheet.  need to set the grid on the update request */
		startGrid.setSheetId(this.getSheetId(GAMEPLAN_SHEET));

		updateCellsRequestList.add(new Request().setUpdateCells(body));               

		BatchUpdateSpreadsheetRequest batchUpdateR = new BatchUpdateSpreadsheetRequest();
		batchUpdateR.setRequests(updateCellsRequestList);
		service.spreadsheets().batchUpdate(this.getSheetId(), batchUpdateR).execute();
		
		
		AutoResizeDimensionsRequest resizeR = new AutoResizeDimensionsRequest ();
	    DimensionRange dimensions = new DimensionRange().setDimension("COLUMNS").setStartIndex(0).setEndIndex(1);
	    dimensions.setSheetId(this.getSheetId(GAMEPLAN_SHEET));
		resizeR.setDimensions(dimensions);
		List<Request> resizeCellsRequestList = new ArrayList<Request>();
		resizeCellsRequestList.add(new Request().setAutoResizeDimensions(resizeR));
	    DimensionRange dimensions2 = new DimensionRange().setDimension("COLUMNS").setStartIndex(2).setEndIndex(7);
	    dimensions2.setSheetId(this.getSheetId(GAMEPLAN_SHEET));
	    AutoResizeDimensionsRequest resizeR2 = new AutoResizeDimensionsRequest ();
		resizeR2.setDimensions(dimensions2);
	    resizeCellsRequestList.add(new Request().setAutoResizeDimensions(resizeR2));
	    
	    DimensionRange dimensions3 = new DimensionRange().setDimension("COLUMNS").setStartIndex(8).setEndIndex(11);
	    dimensions3.setSheetId(this.getSheetId(GAMEPLAN_SHEET));
	    AutoResizeDimensionsRequest resizeR3 = new AutoResizeDimensionsRequest ();
		resizeR3.setDimensions(dimensions3);
	    resizeCellsRequestList.add(new Request().setAutoResizeDimensions(resizeR3));

	    Request nameWidthR = new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                        .setRange(new DimensionRange().setSheetId(this.getSheetId(GAMEPLAN_SHEET)).setDimension("COLUMNS").setStartIndex(1).setEndIndex(2))
                        .setProperties(new DimensionProperties().setPixelSize(200))
                        .setFields("pixelSize"));
	    resizeCellsRequestList.add(nameWidthR); 
	    
	    Request recWidthR = new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                .setRange(new DimensionRange().setSheetId(this.getSheetId(GAMEPLAN_SHEET)).setDimension("COLUMNS").setStartIndex(5).setEndIndex(6))
                .setProperties(new DimensionProperties().setPixelSize(350))
                .setFields("pixelSize"));
	    resizeCellsRequestList.add(recWidthR); 
	    
	    Request notesWidthR = new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                .setRange(new DimensionRange().setSheetId(this.getSheetId(GAMEPLAN_SHEET)).setDimension("COLUMNS").setStartIndex(7).setEndIndex(8))
                .setProperties(new DimensionProperties().setPixelSize(300))
                .setFields("pixelSize"));
	    resizeCellsRequestList.add(notesWidthR); 
	    Request wiWidthR = new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                .setRange(new DimensionRange().setSheetId(this.getSheetId(GAMEPLAN_SHEET)).setDimension("COLUMNS").setStartIndex(9).setEndIndex(12))
                .setProperties(new DimensionProperties().setPixelSize(80))
                .setFields("pixelSize"));
	    resizeCellsRequestList.add(wiWidthR); 
	    
	 	//pushing everything to the sheet
		batchUpdateR.setRequests(resizeCellsRequestList);
		service.spreadsheets().batchUpdate(this.getSheetId(), batchUpdateR).execute();
		
		return;
	}
	private List<CellData> writeRosterRow(Boolean atNewWeight, Boolean oddWeight,Boolean defaultFont,
				String weightData, 
				String nameData, String gradeData, String vjvData, String recordData, 
				String breakdownData, String lastYearData,String notesData, String certWeightData, String seedWeightData) {
	
	    TextFormat f = new TextFormat();
	    f.setFontSize(25);
	    f.setBold(true);
		/* Row Borders */
		Borders rowMidBorders = new Borders();
		rowMidBorders.setBottom(new Border().setStyle("SOLID"));
		Borders rowLeftBorders = new Borders();
		rowLeftBorders.setBottom(new Border().setStyle("SOLID"));
		rowLeftBorders.setLeft(new Border().setStyle("SOLID_THICK"));
		Borders rowRightBorders = new Borders();
		rowRightBorders.setBottom(new Border().setStyle("SOLID"));
		rowRightBorders.setRight(new Border().setStyle("SOLID_THICK"));
		
		Borders nwMidBorders = rowMidBorders.clone(); nwMidBorders.setTop(new Border().setStyle("SOLID_THICK"));
		Borders nwLeftBorders = rowLeftBorders.clone(); nwLeftBorders.setTop(new Border().setStyle("SOLID_THICK"));
		Borders nwRightBorders = rowRightBorders.clone(); nwRightBorders.setTop(new Border().setStyle("SOLID_THICK"));
	
		List<CellData> row = new ArrayList<CellData>();			

		CellData weightDataCell = new CellData(); 
		CellFormat weightDataCellFormat = new CellFormat();
		
		if ( atNewWeight) {
			weightDataCellFormat.setBorders(nwLeftBorders);
		} else {
			weightDataCellFormat.setBorders(rowLeftBorders);
		}
		weightDataCell.setUserEnteredFormat(weightDataCellFormat);
		if ( atNewWeight ) {
			weightDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(weightData)) ;
		} 

		if ( oddWeight) { 
			weightDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(weightDataCell);
	
		CellData nameDataCell = new CellData(); 
		CellFormat nameDataCellFormat = new CellFormat();
		if ( atNewWeight) {
			nameDataCellFormat.setBorders(nwMidBorders);
		} else {
			nameDataCellFormat.setBorders(rowMidBorders);
		}
		nameDataCell.setUserEnteredFormat(nameDataCellFormat);
		if ( oddWeight) { 
			nameDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(nameDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(nameData)) ) ;
		
		System.out.println("working on " + nameData);

		CellData gradeDataCell = new CellData(); 
		CellFormat gradeDataCellFormat = new CellFormat();
		if ( atNewWeight) {
			gradeDataCellFormat.setBorders(nwMidBorders);
		} else {
			gradeDataCellFormat.setBorders(rowMidBorders);
		}
		gradeDataCell.setUserEnteredFormat(gradeDataCellFormat);
		if ( oddWeight) { 
			gradeDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(gradeDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(gradeData)) ) ;
	
	
		CellData jvvDataCell = new CellData(); 
		CellFormat jvvDataCellFormat = new CellFormat();
		if ( atNewWeight) {
			jvvDataCellFormat.setBorders(nwMidBorders);
		} else {
			jvvDataCellFormat.setBorders(rowMidBorders);
		}
		jvvDataCell.setUserEnteredFormat(jvvDataCellFormat);
		if ( oddWeight) { 
			jvvDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(jvvDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(vjvData)))  ;

		CellData recordDataCell = new CellData(); 
		CellFormat recordDataCellFormat = new CellFormat();
		if ( atNewWeight) {
			recordDataCellFormat.setBorders(nwMidBorders);
		} else {
			recordDataCellFormat.setBorders(rowMidBorders);
		}
		recordDataCell.setUserEnteredFormat(gradeDataCellFormat);
		if ( oddWeight) { 
			recordDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(recordDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(recordData)) ) ;

		
		CellData sRecordDataCell = new CellData(); 
		CellFormat sRecordDataCellFormat = new CellFormat();
		if ( atNewWeight) {
			sRecordDataCellFormat.setBorders(nwMidBorders);
		} else {
			sRecordDataCellFormat.setBorders(rowMidBorders);
		}
		sRecordDataCellFormat.setWrapStrategy("WRAP");	
		sRecordDataCell.setUserEnteredFormat(sRecordDataCellFormat);
		if ( oddWeight) { 
			sRecordDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(sRecordDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(breakdownData)))  ;
	
		
		CellData lyRecordDataCell = new CellData(); 
		CellFormat lyRecordDataCellFormat = new CellFormat();
		if ( atNewWeight) {
			lyRecordDataCellFormat.setBorders(nwMidBorders);
		} else {
			lyRecordDataCellFormat.setBorders(rowMidBorders);
		}
		lyRecordDataCell.setUserEnteredFormat(lyRecordDataCellFormat);
		if ( oddWeight) { 
			lyRecordDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(lyRecordDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(lastYearData)) ) ;

		CellFormat notesDataCellFormat = new CellFormat(); 
		notesDataCellFormat.setWrapStrategy("WRAP");				
		if ( defaultFont ) {
			// font set to red.
			f = new TextFormat();
			f.setForegroundColor(redColor);
			notesDataCellFormat.setTextFormat(f);
		} else {
			//font set to blue.
			f = new TextFormat();
			f.setForegroundColor(blueColor);
			notesDataCellFormat.setTextFormat(f);
		}

		CellData notesDataCell = new CellData(); 
		if ( atNewWeight) {
			notesDataCellFormat.setBorders(nwMidBorders);
		} else {
			notesDataCellFormat.setBorders(rowMidBorders);
		}
		notesDataCell.setUserEnteredFormat(notesDataCellFormat);
		if ( oddWeight) { 
			notesDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(notesDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(notesData)) ); 
	
		
		CellData iWIDataCell = new CellData(); 
		CellFormat iWIDataCellFormat = new CellFormat();
		if ( atNewWeight) {
			iWIDataCellFormat.setBorders(nwMidBorders);
		} else {
			iWIDataCellFormat.setBorders(rowMidBorders);
		}
		iWIDataCell.setUserEnteredFormat(iWIDataCellFormat);
		if ( oddWeight) { 
			iWIDataCellFormat.setBackgroundColor(greyColor);
		}
		row.add(iWIDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(certWeightData)) ) ;

		CellData iSeedCell = new CellData(); 
		CellFormat iSeedCellFormat = new CellFormat();
		if ( atNewWeight) {
			iSeedCellFormat.setBorders(nwRightBorders);
		} else {
			iSeedCellFormat.setBorders(rowRightBorders);
		}
		iSeedCell.setUserEnteredFormat(iSeedCellFormat);
		if ( oddWeight) { 
			iSeedCellFormat.setBackgroundColor(greyColor);
		}
		row.add(iSeedCell.setUserEnteredValue(new ExtendedValue().setStringValue(seedWeightData)) ) ;
	
		atNewWeight=false;
		
		return row;
	}
	public void writeVerboseTeam(GSTeam oppTeam,GSTeam homeTeam) throws Exception {
		
		/* If GamePlan exists, delete it and start fresh. */
		if ( this.getSheetId(VERBOSE_TEAM_SHEET) > 0 )  {
			verboseMessage("Verbose Team Sheet found at index " + this.getSheetId(VERBOSE_TEAM_SHEET) );
			List<Request> delRequests = new ArrayList<Request>();
			DeleteSheetRequest delSheet = new DeleteSheetRequest();
			delSheet.setSheetId(this.getSheetId(VERBOSE_TEAM_SHEET));
			    delRequests.add(new Request().setDeleteSheet(delSheet));
				BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
				requestBody.setRequests(delRequests);
				service.spreadsheets().batchUpdate(this.getSheetId(), requestBody).execute();
				verboseMessage("Verbose Team Sheet deleted." );
		}
		
		
		GridCoordinate startGrid = new GridCoordinate().setSheetId(this.getSheetId(VERBOSE_TEAM_SHEET)).setRowIndex(0).setColumnIndex(0);
		UpdateCellsRequest body = new UpdateCellsRequest().setStart(startGrid).setFields("*");
		List<RowData> rows = new ArrayList<RowData>();

		List<CellData> r1 = new ArrayList<CellData>();
		r1.add(new CellData());
		CellData titleCell =new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(homeTeam.getTeamName() + " vs. " + oppTeam.getTeamName() + " Results Details"));
		r1.add(titleCell);
	    CellFormat titleFormat = new CellFormat();
	    titleCell.setUserEnteredFormat(titleFormat);
	    TextFormat f = new TextFormat();
	    f.setFontSize(25);
	    f.setBold(true);
	    titleFormat.setTextFormat(f);
	
		rows.add(new RowData().setValues(r1));

		List<CellData> blankrow = new ArrayList<CellData>();
		rows.add(new RowData().setValues(blankrow));
		rows.add(new RowData().setValues(blankrow));
		
		Hashtable<String,GSWrestler> theRoster = oppTeam.getRoster();
		
		Collection<GSWrestler> wc = theRoster.values();
		ArrayList<GSWrestler> aw = new ArrayList<GSWrestler> (wc);

		Comparator<GSWrestler> compareIt = new Comparator<GSWrestler>() {
			
			public int compare(GSWrestler w1, GSWrestler w2) {
				
				int wS= w1.getFinWeight();
				int w2S= w2.getFinWeight();
				
				
				int i = wS - w2S;
			
				if ( i == 0 ) {
					int w1jvv=0;
					int w2jvv=0;
					if ( w1.getJVOrVarsity().equals(GS_JVV_VARSITY)) {
						w1jvv=4;
					} else if ( w1.getJVOrVarsity().contentEquals(GS_JVV_JV)) {
						w1jvv=3;
					} else if ( w1.getJVOrVarsity().contentEquals(GS_JVV_EXHIBITION)) {
						w1jvv=2;
					} else if ( w1.getJVOrVarsity().contentEquals(GS_JVV_NOSEED)) {
						w1jvv=1;
					} 
					if ( w2.getJVOrVarsity().equals(GS_JVV_VARSITY)) {
						w2jvv=4;
					} else if ( w2.getJVOrVarsity().contentEquals(GS_JVV_JV)) {
						w2jvv=3;
					} else if ( w2.getJVOrVarsity().contentEquals(GS_JVV_EXHIBITION)) {
						w2jvv=2;
					} else if ( w2.getJVOrVarsity().contentEquals(GS_JVV_NOSEED)) {
						w2jvv=1;
					} 
					i = w2jvv - w1jvv;
					if ( i == 0 ) {
						if ( ! w2.getGradeString().equals("") && ! w1.getGradeString().equals("") ) {
							i = Integer.parseInt(w2.getGradeString()) - Integer.parseInt(w1.getGradeString());
					
						}
					}
				}
				return i;
			}
		};
			
		Collections.sort(aw,compareIt);
		
		/* Setup the header style */

		CellFormat headerFormat = new CellFormat();
		TextFormat hf = new TextFormat();
		hf.setBold(true);
		headerFormat.setTextFormat(hf);		
		headerFormat.setWrapStrategy("WRAP");

		Borders hBorders = new Borders();
		hBorders.setTop(new Border().setStyle("SOLID_THICK"));
		hBorders.setBottom(new Border().setStyle("SOLID_THICK"));
		headerFormat.setBorders(hBorders);
		
		List<CellData> headerrow = new ArrayList<CellData>();
		rows.add(new RowData().setValues(headerrow));
		
		CellData weightH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Weight"));
		headerrow.add(weightH);
		CellFormat leftHeaderFormat = headerFormat.clone();
		leftHeaderFormat.getBorders().setLeft(new Border().setStyle("SOLID_THICK"));
		leftHeaderFormat.setWrapStrategy("WRAP");
		weightH.setUserEnteredFormat(leftHeaderFormat);
	
		CellData nameH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Name"));
		headerrow.add(nameH);
		nameH.setUserEnteredFormat(headerFormat);
		
		CellData breakdownH = new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Breakdown"));
		headerrow.add(breakdownH);
		CellFormat rightHeaderFormat = headerFormat.clone();
		rightHeaderFormat.getBorders().setRight(new Border().setStyle("SOLID_THICK"));
		rightHeaderFormat.setWrapStrategy("WRAP");
		breakdownH.setUserEnteredFormat(rightHeaderFormat);
		
		/* Row Borders */
		Borders rowMidBorders = new Borders();
		rowMidBorders.setBottom(new Border().setStyle("SOLID"));
		Borders rowLeftBorders = new Borders();
		rowLeftBorders.setBottom(new Border().setStyle("SOLID"));
		rowLeftBorders.setLeft(new Border().setStyle("SOLID_THICK"));
		Borders rowRightBorders = new Borders();
		rowRightBorders.setBottom(new Border().setStyle("SOLID"));
		rowRightBorders.setRight(new Border().setStyle("SOLID_THICK"));
		Borders rowMidBottomBorders = rowMidBorders.clone(); rowMidBottomBorders.setBottom(new Border().setStyle("SOLID_THICK"));
		Borders rowLeftBottomBorders = rowLeftBorders.clone(); rowLeftBottomBorders.setBottom(new Border().setStyle("SOLID_THICK"));
		Borders rowRightBottomBorders = rowRightBorders.clone(); rowRightBottomBorders.setBottom(new Border().setStyle("SOLID_THICK"));
		
		Borders nwMidBorders = rowMidBorders.clone(); nwMidBorders.setTop(new Border().setStyle("SOLID_THICK"));
		Borders nwLeftBorders = rowLeftBorders.clone(); nwLeftBorders.setTop(new Border().setStyle("SOLID_THICK"));
		Borders nwRightBorders = rowRightBorders.clone(); nwRightBorders.setTop(new Border().setStyle("SOLID_THICK"));
		int count=0;

		while ( aw.size() > count ) {
		
			GSWrestler w = aw.get(count);
		
			List<CellData> row = new ArrayList<CellData>();			
			
			CellData weightDataCell = new CellData(); 
			CellFormat weightDataCellFormat = leftHeaderFormat.clone();
			weightDataCell.setUserEnteredFormat(weightDataCellFormat);
			weightDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(String.valueOf(w.getFinWeight()))) ;
			weightDataCellFormat.setBackgroundColor(this.greyColor);
			row.add(weightDataCell);
			
			CellData nameDataCell = new CellData(); 
			CellFormat nameDataCellFormat = headerFormat.clone();
			nameDataCell.setUserEnteredFormat(nameDataCellFormat);
			nameDataCellFormat.setBackgroundColor(greyColor);
			row.add(nameDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(w.getName())) ) ;
				
			System.out.println("Verbose Team working on " + w.getName());
			
			String ss="";
			//works
			if ( w.getRecordBreakdown().length() > 0 ) {
				ss = w.getRecordBreakdown() + ";" + w.getMatchesAtWeightString();
				ss+= "\n";
			}
			ss+= "Grade: " + w.getGradeString();
			ss+= "\nLast Year: " + w.getTrackLastYearRecord();
			ss+= "\nV/J: " + w.getJVOrVarsity();
			String wIStr="\nCert Weight: ";
			wIStr +=  w.getCert();
			wIStr += "\nSeed Weight: ";
			wIStr += w.getSeed();
			
			CellData breakdownDataCell = new CellData(); 
			CellFormat breakdownDataCellFormat = rightHeaderFormat.clone();
			breakdownDataCell.setUserEnteredFormat(breakdownDataCellFormat);
			breakdownDataCellFormat.setBackgroundColor(greyColor);
			breakdownDataCellFormat.getTextFormat().setBold(false);		
			breakdownDataCellFormat.setWrapStrategy("WRAP");
			row.add(breakdownDataCell.setUserEnteredValue(new ExtendedValue().setStringValue(ss + wIStr)) ) ;

			rows.add(new RowData().setValues(row));

			List<Bout> bouts = w.getBouts();
			boolean h2h = false;
			boolean co = false;
			
			if ( bouts != null ) {
				boolean lastBout=false;
				for ( int i=0; i < bouts.size(); i++ ) {
				    
					if ( i == bouts.size()-1 ) { lastBout = true; }
				
					h2h=false;
					co = false;
					Bout b = bouts.get(i);
					
					row = new ArrayList<CellData>();			
					
					CellData blankWeight = new CellData();
					CellData boutHighlight = new CellData();
					CellData boutDetails = new CellData();
					
					Borders uMidB = rowMidBorders;
					Borders uLeftB = rowLeftBorders;
					Borders uRightB = rowRightBorders;
					if ( lastBout ) {
						uMidB = rowMidBottomBorders;
						uLeftB = rowLeftBottomBorders;
						uRightB = rowRightBottomBorders;
					} 
				
					CellFormat blankWeightFormat = leftHeaderFormat.clone();
					blankWeight.setUserEnteredFormat(blankWeightFormat);
					blankWeight.setUserEnteredValue(new ExtendedValue().setStringValue("")) ;
					blankWeightFormat.setBorders(uLeftB);
					row.add(blankWeight);
				
					String sM = b.toString();
				
					if ( b.getOpponentTeam().equals(homeTeam.getTeamName())) {
						h2h=true;
						sM += "\n---- HEAD TO HEAD ----";
					}
					String k = b.getOpponentTeam() + ":" + b.getOpponentName();
					List<Bout> bb = homeTeam.lookupCommonMatch(k);
					if ( bb != null ) {
					   if ( bb.size() > 0 ) {
						   co = true;
						   sM += "\n---- COMMON ----";
					   }
					   for ( int ii=0; ii < bb.size(); ii++ ) {
						  Bout commonMatch = bb.get(ii);
						  sM+= "\n" + commonMatch.toString();
					   }
					}
				
					String msg="";
					CellFormat boutHighlightFormat = headerFormat.clone();
					boutHighlight.setUserEnteredFormat(boutHighlightFormat);

					if ( b.isLossByInjury() ) {
						// font set to red.
					    f = new TextFormat();
					    f.setForegroundColor(redColor);
					    boutHighlightFormat.setTextFormat(f);
					    boutHighlightFormat.getTextFormat().setBold(true);		
						msg="Injury Default";
					} else if ( h2h == true | co == true ) {
						// font set to blue.
					    f = new TextFormat();
					    f.setForegroundColor(blueColor);
					    boutHighlightFormat.setTextFormat(f);
					    boutHighlightFormat.getTextFormat().setBold(true);
						if (h2h) { msg = "Head to Head"; } else { msg="Common";}
						
					} 
					
					boutHighlight.setUserEnteredValue(new ExtendedValue().setStringValue(msg)) ;
					boutHighlightFormat.setBorders(uMidB);
					row.add(boutHighlight);
					
					CellFormat boutDetailsFormat = headerFormat.clone();
					boutDetails.setUserEnteredFormat(boutDetailsFormat);
					boutDetails.setUserEnteredValue(new ExtendedValue().setStringValue(sM)) ;
					boutDetailsFormat.getTextFormat().setBold(false);		

					boutDetailsFormat.setBorders(uRightB);
					row.add(boutDetails);
					
					rows.add(new RowData().setValues(row));

				}
			}
			// insert blank row after each wrestler.
			row = new ArrayList<CellData>();	
			CellData b = new CellData();
			b.setUserEnteredValue(new ExtendedValue().setStringValue("")) ;
			row.add(b);
			rows.add(new RowData().setValues(row));
			
			count++;
			
		}

		AddSheetRequest addSheet = new AddSheetRequest();
		addSheet.setProperties(new SheetProperties().setTitle(VERBOSE_TEAM_SHEET));
		List<Request> requests = new ArrayList<Request>();
		    
		requests.add(new Request().setAddSheet(addSheet));
		BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
		requestBody.setRequests(requests);

		service.spreadsheets().batchUpdate(this.getSheetId(), requestBody).execute();
				
		/* Setting the body data of the gameplan sheet. */
		body.setRows(rows);
		List<Request> updateCellsRequestList = new ArrayList<Request>();

		/* We created a new gameplan sheet.  need to set the grid on the update request */
		startGrid.setSheetId(this.getSheetId(VERBOSE_TEAM_SHEET));

		updateCellsRequestList.add(new Request().setUpdateCells(body));               

		BatchUpdateSpreadsheetRequest batchUpdateR = new BatchUpdateSpreadsheetRequest();
		batchUpdateR.setRequests(updateCellsRequestList);
		service.spreadsheets().batchUpdate(this.getSheetId(), batchUpdateR).execute();
		
		AutoResizeDimensionsRequest resizeR = new AutoResizeDimensionsRequest ();
	    DimensionRange dimensions = new DimensionRange().setDimension("COLUMNS").setStartIndex(0).setEndIndex(1);
	    dimensions.setSheetId(this.getSheetId(VERBOSE_TEAM_SHEET));
		resizeR.setDimensions(dimensions);
		List<Request> resizeCellsRequestList = new ArrayList<Request>();
		resizeCellsRequestList.add(new Request().setAutoResizeDimensions(resizeR));
		
		Request nameWidthR = new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                   .setRange(new DimensionRange().setSheetId(this.getSheetId(VERBOSE_TEAM_SHEET)).setDimension("COLUMNS").setStartIndex(1).setEndIndex(2))
                   .setProperties(new DimensionProperties().setPixelSize(200))
                   .setFields("pixelSize"));
		resizeCellsRequestList.add(nameWidthR); 

	    Request breakdownWidthR = new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                        .setRange(new DimensionRange().setSheetId(this.getSheetId(VERBOSE_TEAM_SHEET)).setDimension("COLUMNS").setStartIndex(2).setEndIndex(3))
                        .setProperties(new DimensionProperties().setPixelSize(800))
                        .setFields("pixelSize"));
	    resizeCellsRequestList.add(breakdownWidthR); 
	 	    
		batchUpdateR.setRequests(resizeCellsRequestList);
		service.spreadsheets().batchUpdate(this.getSheetId(), batchUpdateR).execute();
		verboseMessage("Leaving Verbose Team Write");
		return;
	}
    public void writeTeam(GSTeam theOppTeam,GSTeam homeTeam) {
        
		verboseMessage("Working with sheet <" + this.getSheetId() + "> team <" + this.getTeam() + ">"); 

    	try {
			this.writeRoster(theOppTeam,homeTeam);
			this.writeVerboseTeam(theOppTeam,homeTeam);
		} catch (Exception e ) {
			System.out.println("crap");
			e.printStackTrace();
		}
		return;
		
	}
}