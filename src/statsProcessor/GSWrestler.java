package statsProcessor;

/*
 *
 *  This is a class to holds an wrestler and all data pertaining to him/her.
 */
class GSWrestler extends Wrestler {
	
	private String seed="";
	private String cert="";
	private String jVOrVarsity="";
	
	
	public GSWrestler(String n, String team){
		super(n,team);
		return;
	}

	public void setJVOrVarsity(String jvv) {
		this.jVOrVarsity=jvv;
	}
	public String getJVOrVarsity() {
		return this.jVOrVarsity;
	}

	
    public void printVerbose () {
		System.out.println( getName() + 
			" Grade<" + getGrade() + ">" + 
			" Seed<" + getSeed() + ">" +
			" Cert<" + getCert() + ">"); 

    }	  

	public String getSeed() { return seed; }
	public String getCert() { return cert; }
	
	public String getGradeString() {
		if ( this.getGrade() == WrestlingLanguage.Grade.G5 ) {
			return "5";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.G6 ) {
			return "6";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.G7 ) {
			return "7";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.G8 ) {
			return "8";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.GK ) {
			return "0";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.G1 ) {
			return "1";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.G2 ) {
			return "2";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.G3 ) {
			return "3";
		} else if ( this.getGrade() == WrestlingLanguage.Grade.G4 ) {
			return "4";
		}
		return "";
	}

	public void setSeed( String n ) {
		seed = n.strip();
	}
	public void setCert( String n ) {
		cert = n.strip();
	}
	
	public void setGrade5() { this.setGrade(WrestlingLanguage.Grade.G5);	}
	public void setGrade6() { this.setGrade(WrestlingLanguage.Grade.G6);	}
	public void setGrade7() { this.setGrade(WrestlingLanguage.Grade.G7);	}
	public void setGrade8() { this.setGrade(WrestlingLanguage.Grade.G8);	}
	
	
	public String toString() {
		return super.toString();
	}

	public int getFinWeight() {
		
		int certLookup=this.getLowestCertWeight();
		
		if ( certLookup > 0 ) { 
			return certLookup;
		}
		int finWeight = 9999;
		finWeight = Integer.parseInt( this.getLowestWeightWrestled());

		if ( finWeight == 9999 ) {
			if ( ! seed.equals("") ) {
				finWeight = Integer.parseInt(seed);
			}
		}
		if ( finWeight == 0) { 
			finWeight=9999; 
		}
		return finWeight;
	}
	private int getLowestCertWeight() {
		if ( this.cert.equals("") || this.cert.equals("0")) {
			return 0;
		}
		Double c=Double.parseDouble(this.cert);
		int[] weights=GSDualMeet.getWeightListInt();
		
		for ( var i=0; i < weights.length; i++ ) {
			Double at = weights[i] + 0.2 ;
			if ( c <= at ) {
				return weights[i];
			}
		}
		return weights[weights.length-1];
	}


}
