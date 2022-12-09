package statsProcessor;


/**
 * This is the main factory that extracts from track.
 * There will be a couple of extraction options.  
 * 1) Extract from a copy in Excel.
 * 2) Extract using screen scraping.
 * 3) Extract using api calls.  I'm not sure this will ever be possible.
 */
class ExtractFactory {
   
	private static String fileType; 		//type of file google or excel
	private static String opponentSheetId;  //opponent sheet id if using google.
    private static String opponentFile;		//opponent file if using excel.
	private static String opponentTeam="";  // opponent team name ( this is the name the website uses for them.
	private static String homeSheetId;		// home sheet id if using google.
	private static String homeFile;			// home file if using excel.
	private static String homeTeam="";		// home team name ( this is the name the website uses for them.
	private static String CJResultsId;			// home file if using excel.
	private static String ageLevel=""; 		//This should be GS or HS
	
    private static void processArgs(String[] args) throws Exception {
		int i=0;
		System.out.println("args are " + args.length + "    ->" + args );
		while ( i < args.length) {
			System.out.println( "arg of " + i + "=<" + args[i] + ">");
			i++;
		}
		i=0;
        while ( i < args.length ) {
			String arg=args[i]; 
		System.out.println("Working on <" + arg + "> i=" + i);
			if (arg.equals("-of") ) {
				i++;			 
				opponentFile=args[i];
			} else if (arg.equals("-ot") ) {
				i++;
				opponentTeam = args[i];
			
			} else if (arg.equals("-ht") ) {
				i++;
				homeTeam = args[i];
			} else if (arg.equals("-hf") ) {
				i++;
				homeFile = args[i];
			} else if (arg.equals("-hsId") ) {
				i++;
				homeSheetId = args[i];
			} else if (arg.equals("-osId") ) {
				i++;
				opponentSheetId = args[i];
			} else if (arg.equals("-cjId") ) {
				
				i++;
				CJResultsId= args[i];
			} else if (arg.equals("-type") ) {
				i++;
				fileType = args[i];
			} else if (arg.equals("-al") ) {
				i++;
				ageLevel=args[i];
			} else {
				 throw new Exception ("Unknown Argument <" + args[i]);
			}
			if ( ! fileType.equals("google") && ! fileType.equals("excel")) {
				throw new Exception("-type needs to be set to excel | google.");
			}
		    i++;
	    }
	}
    public static void main(String[] args) {    	

        System.out.println("Starting..."); // Display the string.
        System.out.println("New project..."); // Display the string.
        
        try {
			processArgs(args);
			if ( fileType.equals("google")) {
				
				System.out.println("Opponent =<" + opponentTeam + "> opponentSheetId=<" + opponentSheetId + "> homeTeam<" + homeTeam + "> homeSheetId=<" + homeSheetId + ">" + "cjId>"+ CJResultsId);

				if (ageLevel.equals("GS") ) {
					GSGoogleSheetsExtractor extractor = new GSGoogleSheetsExtractor(opponentTeam,opponentSheetId);
					//use for all teams
					//GoogleSheetsExtractor extractor = new GoogleSheetsExtractor(CJResultsId);
					System.out.println("Working with file <" + opponentSheetId + "> opponentTeam <" + opponentTeam + ">"); 
					GSTeam theOpponentTeam = extractor.extractTeam();
					GSGoogleSheetsExtractor homeExtractor = new GSGoogleSheetsExtractor(homeTeam,homeSheetId);
					GSTeam homeTeam= homeExtractor.extractTeam();
					GSGoogleSheetsWriter e = new GSGoogleSheetsWriter(opponentTeam,opponentSheetId);
					e.writeTeam(theOpponentTeam,homeTeam);
				} else if ( ageLevel.contentEquals("HS" ) ) {
					HSGoogleSheetsExtractor extractor = new HSGoogleSheetsExtractor(opponentTeam,opponentSheetId);
					//use for all teams
					//GoogleSheetsExtractor extractor = new GoogleSheetsExtractor(CJResultsId);
					System.out.println("Working with file <" + opponentSheetId + "> opponentTeam <" + opponentTeam + ">"); 
					Team theOpponentTeam = extractor.extractTeam();
					HSGoogleSheetsExtractor homeExtractor = new HSGoogleSheetsExtractor(homeTeam,homeSheetId);
					Team homeTeam= homeExtractor.extractTeam();
					HSGoogleSheetsWriter e = new HSGoogleSheetsWriter(opponentTeam,opponentSheetId);
					e.writeTeam(theOpponentTeam,homeTeam);
				} else {
					System.out.println("You need to set -al for Age Level.  Values are 'GS' or 'HS'.");
				}
			} else if ( fileType.equals("excel")) {
				System.out.println("Opponent =<" + opponentTeam + "> opponentFile=<" + opponentFile + "> homeTeam<" + homeTeam + "> homeFile=<" + homeFile + ">");
				ExcelExtractor extractor = new ExcelExtractor(opponentTeam,opponentFile);
				System.out.println("Working with file <" + opponentFile + "> opponentTeam <" + opponentTeam + ">"); 
				GSTeam theOpponentTeam = extractor.extractTeam();
				ExcelExtractor homeExtractor = new ExcelExtractor(homeTeam,homeFile);
				GSTeam homeTeam= homeExtractor.extractTeam();
				ExcelWriter e = new ExcelWriter(opponentTeam,opponentFile);
				e.writeTeam(theOpponentTeam,homeTeam);
				
			} else {
				System.out.println("FileType wrong or not set!!!");
			}

			System.out.println("Done");
		} catch (Exception e ) {
			System.out.println("crap");
			e.printStackTrace();
		}
		
	}
}
