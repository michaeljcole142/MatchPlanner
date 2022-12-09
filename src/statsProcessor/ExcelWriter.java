package statsProcessor;


import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * This class is used to write to excel any of the data that has been
 * accumulated.
 */
class ExcelWriter  {
 
  
	private String filename;
	private String team;
	private boolean verbose=true;
	private boolean overwrite=true;
	public static String GAMEPLAN_SHEET="Gameplan";
	public static String VERBOSE_TEAM_SHEET="VerboseTeam";
	
    public ExcelWriter(String t, String f) {
		setTeam(t);
		setFileName(f);
	}
	public String getFileName() { return filename; }
	public String getTeam() { return team; }
	
	public void setFileName(String f) { filename = f; }
	public void setTeam(String t) { team = t; }
	
	public boolean isOverwriteOn() {
		return overwrite;
	}
	
	/* */
	private void verboseMessage(String m ) {
		if ( verbose ) {
			System.out.println(m);
		}
	}
	public void writeRoster(Sheet rSheet, GSTeam oppTeam,GSTeam homeTeam) {
	
		Workbook wb = rSheet.getWorkbook();
		
		rSheet.setMargin(Sheet.RightMargin, 0.25);
		rSheet.setMargin(Sheet.LeftMargin, 0.25);
		rSheet.setMargin(Sheet.TopMargin, 0.5);
		rSheet.setMargin(Sheet.BottomMargin, 0.5);
		
		Hashtable<String,GSWrestler> theRoster = oppTeam.getRoster();
				
		Collection<GSWrestler> wc = theRoster.values();
		ArrayList<GSWrestler> aw = new ArrayList<GSWrestler> (wc);
		
		Comparator<GSWrestler> compareIt = new Comparator<GSWrestler>() {
			public int compare(GSWrestler w1, GSWrestler w2) {
				String wS=w1.getPrintWeight();
				String w2S=w2.getPrintWeight();
				
				int i = wS.compareTo(w2S);
				if ( i == 0 ) {
					int m1 = w1.getPrintRank();
					int m2 = w2.getPrintRank();
					i=m2-m1;
				}
				return i;
			}
		};
		Collections.sort(aw,compareIt);
		
		CellStyle  header = wb.createCellStyle();
		Font bF = wb.createFont();
		bF.setFontHeightInPoints((short) 25);
		bF.setBold(true);
		Row rr = rSheet.createRow(0);
		Cell cc = rr.createCell(1);
		header.setFont(bF);
		CellRangeAddress region = CellRangeAddress.valueOf("B1:I1");

		// merging the region
		rSheet.addMergedRegion(region);
		cc.setCellStyle(header);
		cc.setCellValue(homeTeam.getTeamName() + " vs. " + oppTeam.getTeamName() + " Overview");
		
		CellStyle hstyle = rSheet.getWorkbook().createCellStyle();

		Font bFont = rSheet.getWorkbook().createFont();
		bFont.setBold(true);
		
		hstyle.setFont(bFont);
		hstyle.setWrapText(true);
		hstyle.setBorderTop(BorderStyle.THICK);
		CellStyle hFirstCellStyle = rSheet.getWorkbook().createCellStyle();
		hFirstCellStyle.cloneStyleFrom(hstyle);
		hFirstCellStyle.setBorderLeft(BorderStyle.THICK);
		hFirstCellStyle.setFont(bFont);
		CellStyle hLastCellStyle = rSheet.getWorkbook().createCellStyle();
		hLastCellStyle.cloneStyleFrom(hstyle);
		hLastCellStyle.setBorderRight(BorderStyle.THICK);
		hLastCellStyle.setFont(bFont);
	
		int sheetRowNum=4;
		Row r = rSheet.createRow(3);
		Cell c = r.createCell(0);
		c.setCellStyle(hFirstCellStyle);
		c.setCellValue("Weight");
		c = r.createCell(1);
		c.setCellStyle(hstyle);
		c.setCellValue("Name");
		c = r.createCell(2);			
		c.setCellStyle(hstyle);
		c.setCellValue("Grade");
		c = r.createCell(3);	
		c.setCellStyle(hstyle);
		c.setCellValue("Record");                                                                                                    
		c = r.createCell(4);
		c.setCellStyle(hstyle);
		c.setCellValue("BreakDown");
		c = r.createCell(5);
		c.setCellStyle(hstyle);
		c.setCellValue("Last Year");
		c = r.createCell(6);
		c.setCellStyle(hstyle);
		c.setCellValue("Prestige");
		c = r.createCell(7);
		c.setCellStyle(hstyle);
		c.setCellValue("Notes");
		c = r.createCell(8);
		c.setCellStyle(hstyle);
		c.setCellValue("Initial WeighIn");
		c = r.createCell(9);
		c.setCellStyle(hstyle);
		c.setCellValue("Last WeighIn Date" );
		c = r.createCell(10);
		c.setCellStyle(hLastCellStyle);
		c.setCellValue("Last WeighIn");	
		
		int count=0;
		String atWeightStr="";
		boolean oddWeight=true;

		boolean atNewWeight=true;
		while ( aw.size() > count ) {
			GSWrestler w = aw.get(count);
			System.out.println("w=" + w);
			
			String wWeight = w.getPrintWeight();
			
			if (! wWeight.equals(atWeightStr)) {
				atNewWeight=true;
				if ( atWeightStr.length() > 0 ) {
					if (oddWeight) {
						oddWeight=false;
					} else {
						oddWeight=true;
					}
				}
			}
		
			atWeightStr = wWeight;
			r = rSheet.createRow(sheetRowNum);
			CellStyle style = rSheet.getWorkbook().createCellStyle();
			
			style.setWrapText(true);
			style.setBorderBottom(BorderStyle.THIN);
			if ( oddWeight ) {
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			}
			if( atNewWeight) {
				style.setBorderTop(BorderStyle.THICK);
			}
			c = r.createCell(0);
			CellStyle firstCellStyle = rSheet.getWorkbook().createCellStyle();
			firstCellStyle.cloneStyleFrom(style);
			firstCellStyle.setBorderLeft(BorderStyle.THICK);
			
			CellStyle lastCellStyle = rSheet.getWorkbook().createCellStyle();
			lastCellStyle.cloneStyleFrom(style);
			lastCellStyle.setBorderRight(BorderStyle.THICK);
			
			c.setCellStyle(firstCellStyle);
			if ( atNewWeight ) {
				c.setCellValue(w.getPrintWeight());
			}
			c = r.createCell(1);
			c.setCellStyle(style);
			c.setCellValue(w.getName());
			System.out.println("working on " + w.getName());
			c = r.createCell(2);	
			c.setCellStyle(style);
WrestlingLanguage.Grade g = w.getGrade();
String gg = "NoGrade";
if (g != null) { gg = g.toString(); } 
			c.setCellValue(gg);			
			c = r.createCell(3);	
			c.setCellStyle(style);
			c.setCellValue(w.getRecordString());
			c = r.createCell(4);
			c.setCellStyle(style);
			if ( w.getRecordBreakdown().length() > 0 ) {
				String ss = w.getRecordBreakdown() + ";" + w.getMatchesAtWeightString();
				c.setCellStyle(style);
				c.setCellValue(ss);
			}
			c = r.createCell(5);
			c.setCellStyle(style);
			c.setCellValue(w.getTrackLastYearRecord());
			c = r.createCell(6);
			c.setCellStyle(style);
			String pres = "";
			if ( w.getPrestigeLastYear() != null ) { pres += w.getPrestigeLastYear().toString() + "(LY)"; }
			if ( w.getPrestige2YearsAgo() != null ) { pres += w.getPrestige2YearsAgo().toString() + "(2YR)"; }
			if ( w.getPrestige3YearsAgo() != null ) { pres += w.getPrestige3YearsAgo().toString() + "(3YR)"; }

			c.setCellValue(pres);
			
			c = r.createCell(7);
			c.setCellStyle(style);

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
			String notes="";
			if ( h2h.length() > 0 ) {
				if (common.length() > 0 ) {
					notes = h2h + "\n" + common;
				} else {
					notes = h2h;
				}
			} else {
				if ( common.length() > 0 ) {
					notes = common;
				}
			}
			
			if ( w.getLossByInjury() ) {
				Font redFont = rSheet.getWorkbook().createFont();
				redFont.setColor(IndexedColors.RED.getIndex());
				CellStyle ss = rSheet.getWorkbook().createCellStyle();
				ss.cloneStyleFrom(style);
				ss.setFont(redFont);

				c.setCellStyle(ss);
				if ( notes.length() > 0 ) {
					notes = notes + "\n" + w.getLossByInjuryString();
				}
				c.setCellValue(notes);
			} else {
				Font blFont = rSheet.getWorkbook().createFont();
				blFont.setColor(IndexedColors.BLUE.getIndex());
				CellStyle ss = rSheet.getWorkbook().createCellStyle();
				ss.cloneStyleFrom(style);
				ss.setFont(blFont);
				c.setCellStyle(ss);
				c.setCellValue(notes);				
			}
			c = r.createCell(8);
			c.setCellStyle(style);
			WeighInHistory wih = w.getWeighInHistory();
			if ( wih == null ) {
				c.setCellValue("None");
				c = r.createCell(9);
				c.setCellStyle(style);
				c = r.createCell(10);
				c.setCellStyle(lastCellStyle);
			} else {
				c.setCellValue(String.format(java.util.Locale.US, "%.1f",w.getWeighInHistory().getInitialWI()));
				c = r.createCell(9);
				c.setCellStyle(style);
				c.setCellValue(wih.getLastWIDate());
				c = r.createCell(10);
				c.setCellStyle(lastCellStyle);
				c.setCellValue(String.format(java.util.Locale.US, "%.1f",w.getWeighInHistory().getLastWI()));

			}
			atNewWeight=false;
			count++;
			sheetRowNum++;
			
		}
		rSheet.autoSizeColumn(0);
		rSheet.autoSizeColumn(1);
		rSheet.autoSizeColumn(2);
		rSheet.autoSizeColumn(3);
		rSheet.setColumnWidth(4,40*256);
		rSheet.autoSizeColumn(5);
		rSheet.autoSizeColumn(6);
		rSheet.setColumnWidth(7,50*256);
		rSheet.setColumnWidth(8,7*256);
		rSheet.setColumnWidth(9,11*256);
		rSheet.setColumnWidth(10,7*256);


		wb.setPrintArea(wb.getSheetIndex(GAMEPLAN_SHEET),0,11,0,sheetRowNum);
		
		rSheet.getPrintSetup().setLandscape(true);
		rSheet.getPrintSetup().setFitWidth(((short)1));
		rSheet.getPrintSetup().setFitHeight(((short)5));
		rSheet.setFitToPage(true);
		
		return;
	}
	public void writeVerboseTeam(Sheet rSheet, GSTeam oppTeam,GSTeam homeTeam) {
		
		Workbook wb = rSheet.getWorkbook();
		
		rSheet.setMargin(Sheet.RightMargin, 0.25);
		rSheet.setMargin(Sheet.LeftMargin, 0.25);
		rSheet.setMargin(Sheet.TopMargin, 0.5);
		rSheet.setMargin(Sheet.BottomMargin, 0.5);
		
		CellStyle  header = wb.createCellStyle();
		Font bF = wb.createFont();
		bF.setFontHeightInPoints((short) 25);
		bF.setBold(true);
		Row rr = rSheet.createRow(0);
		Cell cc = rr.createCell(1);
		header.setFont(bF);
		cc.setCellStyle(header);
		cc.setCellValue(homeTeam.getTeamName() + " vs. " + oppTeam.getTeamName() + " Results Details");
		
		Hashtable<String,GSWrestler> theRoster = oppTeam.getRoster();
		
		Collection<GSWrestler> wc = theRoster.values();
		ArrayList<GSWrestler> aw = new ArrayList<GSWrestler> (wc);

		Comparator<GSWrestler> compareIt = new Comparator<GSWrestler>() {
			public int compare(GSWrestler w1, GSWrestler w2) {
				String wS=w1.getPrintWeight();
				String w2S=w2.getPrintWeight();
				
				int i = wS.compareTo(w2S);
				if ( i == 0 ) {
					int m1 = w1.getPrintRank();
					int m2 = w2.getPrintRank();
					i=m2-m1;
				}
				return i;
			}
		};
		Collections.sort(aw,compareIt);
		
		/* Setup the header style */
		CellStyle hstyle = rSheet.getWorkbook().createCellStyle();
		Font bFont = rSheet.getWorkbook().createFont();
		bFont.setBold(true);
		hstyle.setFont(bFont);
		hstyle.setWrapText(true);
		hstyle.setBorderTop(BorderStyle.THICK);
		CellStyle hFirstCellStyle = rSheet.getWorkbook().createCellStyle();
		hFirstCellStyle.cloneStyleFrom(hstyle);
		hFirstCellStyle.setBorderLeft(BorderStyle.THICK);
		CellStyle hLastCellStyle = rSheet.getWorkbook().createCellStyle();
		hLastCellStyle.cloneStyleFrom(hstyle);
		hLastCellStyle.setBorderRight(BorderStyle.THICK);
	
		
		int sheetRowNum=4;
		Row r = rSheet.createRow(3);
		Cell c = r.createCell(0);
		c.setCellStyle(hFirstCellStyle);
		c.setCellValue("Weight");
		c = r.createCell(1);
		c.setCellStyle(hstyle);
		c.setCellValue("Name");
		c = r.createCell(2);			
		c.setCellStyle(hLastCellStyle);
		c.setCellValue("BreakDown");
		
		
		int count=0;

		while ( aw.size() > count ) {
		
			GSWrestler w = aw.get(count);
		
			r = rSheet.createRow(sheetRowNum);
			
			/* Setup Style for the wrestler */
			CellStyle wStyle = rSheet.getWorkbook().createCellStyle();
			wStyle.setWrapText(true);
			wStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			wStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			wStyle.setBorderTop(BorderStyle.THICK);
			wStyle.setBorderBottom(BorderStyle.THICK);
			CellStyle wFirstCellStyle = rSheet.getWorkbook().createCellStyle();
			wFirstCellStyle.cloneStyleFrom(wStyle);
			wFirstCellStyle.setBorderLeft(BorderStyle.THICK);
			CellStyle wLastCellStyle = rSheet.getWorkbook().createCellStyle();
			wLastCellStyle.cloneStyleFrom(wStyle);
			wLastCellStyle.setBorderRight(BorderStyle.THICK);

			CellStyle bCellStyle = rSheet.getWorkbook().createCellStyle();
			bCellStyle.cloneStyleFrom(wStyle);
			bCellStyle.setFont(bFont);
	
			/* Cell setup for Opponent rows. */
			CellStyle oStyle = rSheet.getWorkbook().createCellStyle();
			oStyle.setWrapText(true);
			oStyle.setBorderBottom(BorderStyle.THIN);
			CellStyle oFirstCellStyle = rSheet.getWorkbook().createCellStyle();
			oFirstCellStyle.cloneStyleFrom(oStyle);
			oFirstCellStyle.setBorderLeft(BorderStyle.THICK);
			CellStyle oLastCellStyle = rSheet.getWorkbook().createCellStyle();
			oLastCellStyle.cloneStyleFrom(oStyle);
			oLastCellStyle.setBorderRight(BorderStyle.THICK);

			
			CellStyle firstCellStyle = rSheet.getWorkbook().createCellStyle();
			firstCellStyle.cloneStyleFrom(wStyle);
			firstCellStyle.setBorderLeft(BorderStyle.THICK);
			CellStyle lastCellStyle = rSheet.getWorkbook().createCellStyle();
			lastCellStyle.cloneStyleFrom(wStyle);
			lastCellStyle.setBorderRight(BorderStyle.THICK);
		
			/* Buildout the wrestler headline row */
			
			c = r.createCell(0);	
			c.setCellStyle(wFirstCellStyle);
			c.setCellValue(w.getPrintWeight());
			c = r.createCell(1);
			c.setCellStyle(bCellStyle);
			c.setCellValue(w.getName());
			c = r.createCell(2);
			c.setCellStyle(wLastCellStyle);
			
			String ss="";
			if ( w.getRecordBreakdown().length() > 0 ) {
				ss = w.getRecordBreakdown() + ";" + w.getMatchesAtWeightString();
				ss+= "\nGrade: " + w.getGrade();
				ss+= "\nLast Year: " + w.getTrackLastYearRecord();
				ss+= "\nPrestige Last Year: " + w.getPrestigeLastYear();
			}
			
			WeighInHistory wih = w.getWeighInHistory();
			String wIStr="\nInitial WeighIn: ";
			if ( wih == null ) {
				wIStr += "None";
				
			} else {
				wIStr +=  String.format(java.util.Locale.US, "%.1f",w.getWeighInHistory().getInitialWI());
				wIStr += " Last WeighIn " + wih.getLastWIDate() + " ";
				wIStr += String.format(java.util.Locale.US, "%.1f",w.getWeighInHistory().getLastWI());
			}
            c.setCellValue(ss + wIStr);
            
			sheetRowNum++;
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
					r = rSheet.createRow(sheetRowNum);
					c = r.createCell(0);
					
					if ( lastBout ) {
						CellStyle ebStyle = rSheet.getWorkbook().createCellStyle();
						ebStyle.cloneStyleFrom(oFirstCellStyle);
						ebStyle.setBorderBottom(BorderStyle.THICK);
					    c.setCellStyle(ebStyle);
					} else {
						c.setCellStyle(oFirstCellStyle);
					}
					/* do cell1 after cell2 */

				    c = r.createCell(2);
				    if ( lastBout ) {
						CellStyle ebStyle = rSheet.getWorkbook().createCellStyle();
						ebStyle.cloneStyleFrom(oLastCellStyle);
						ebStyle.setBorderBottom(BorderStyle.THICK);
					    c.setCellStyle(ebStyle);
					} else {
						c.setCellStyle(oLastCellStyle);
					}
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
					c.setCellValue(sM);

					/* this cell's font is determined by next cell. */
					c = r.createCell(1);
					String msg="";
					if ( b.isLossByInjury() ) {
						Font redFont = rSheet.getWorkbook().createFont();
						redFont.setColor(IndexedColors.RED.getIndex());
						redFont.setBold(true);
						CellStyle roStyle = rSheet.getWorkbook().createCellStyle();
						roStyle.cloneStyleFrom(oStyle);
						roStyle.setFont(redFont);
						if( lastBout ) { roStyle.setBorderBottom(BorderStyle.THICK); }
						
						c.setCellStyle(roStyle);
						msg="Injury Default";
					} else if ( h2h == true | co == true ) {
						Font blueFont = rSheet.getWorkbook().createFont();
						blueFont.setColor(IndexedColors.BLUE.getIndex());
						blueFont.setBold(true);
						CellStyle roStyle = rSheet.getWorkbook().createCellStyle();
						roStyle.cloneStyleFrom(oStyle);
						roStyle.setFont(blueFont);
						if( lastBout ) { roStyle.setBorderBottom(BorderStyle.THICK); }
						c.setCellStyle(roStyle);
						if (h2h) { msg = "Head to Head"; } else { msg="Common";}
						
					} else {
						CellStyle roStyle = rSheet.getWorkbook().createCellStyle();
						roStyle.cloneStyleFrom(oStyle);
						if( lastBout ) { roStyle.setBorderBottom(BorderStyle.THICK); }
						c.setCellStyle(roStyle);
					}
					c.setCellValue(msg);

				    sheetRowNum++;
				}
			}
			count++;
			sheetRowNum++;
			
		}
		
		rSheet.setColumnWidth(2,100*256);
		rSheet.setColumnWidth(1,20*256);
		rSheet.autoSizeColumn(3);
		wb.setPrintArea(wb.getSheetIndex(VERBOSE_TEAM_SHEET),0,2,0,sheetRowNum);
		rSheet.getPrintSetup().setFitWidth(((short)1));
		rSheet.getPrintSetup().setFitHeight(((short)49));
		rSheet.setFitToPage(true);
		return;
	}
    public void writeTeam(GSTeam theOppTeam,GSTeam homeTeam) {

        try {
  	   
			verboseMessage("Working with file <" + filename + "> team <" + team + ">"); 
			Workbook workbook = WorkbookFactory.create(new File(filename));
		  
			verboseMessage("Workbook has " + workbook.getNumberOfSheets());
			verboseMessage("Retrieving Sheets using for-each loop");
			for(Sheet sheet: workbook) {
				verboseMessage("=> " + sheet.getSheetName());
			}
		  
			Sheet gameplanSheet = workbook.getSheet(GAMEPLAN_SHEET);
		  
			if ( gameplanSheet != null && isOverwriteOn() ) {
				int ind = workbook.getSheetIndex(gameplanSheet);
				verboseMessage("Gameplan Sheet found at index " + ind );
				workbook.removeSheetAt(ind);
				verboseMessage("deleted Gameplan sheet.");
			}
			Sheet s = workbook.createSheet();
			workbook.setSheetName(workbook.getSheetIndex(s),GAMEPLAN_SHEET);
			
			writeRoster(s,theOppTeam,homeTeam);

			Sheet verboseTeamSheet = workbook.getSheet(VERBOSE_TEAM_SHEET);
			
			if ( verboseTeamSheet != null && isOverwriteOn() ) {
				int ind = workbook.getSheetIndex(verboseTeamSheet);
				verboseMessage("Verbose Team Sheet found at Index " + ind);
				workbook.removeSheetAt(ind);
				verboseMessage("deleted VerboseTeam Sheet");
			}
			s=workbook.createSheet();
			workbook.setSheetName(workbook.getSheetIndex(s),VERBOSE_TEAM_SHEET);
			
			writeVerboseTeam(s,theOppTeam,homeTeam);
			
			OutputStream fileOut = new FileOutputStream("testwriting.xlsx");
			workbook.write(fileOut);
			
		} catch (Exception e ) {
			System.out.println("crap");
			e.printStackTrace();
		}
		return;
		
	}
}