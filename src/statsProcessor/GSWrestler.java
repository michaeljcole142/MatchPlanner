package statsProcessor;

/*
 *
 *  This is a class to holds an wrestler and all data pertaining to him/her.
 */
class GSWrestler extends Wrestler {
	
	private String seed="";
	private String cert="";
	private String jVOrVarsity;
	
	
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
		int finWeight = 9999;
	
		finWeight = Integer.parseInt( this.getLowestWeightWrestled());

		if ( ! seed.equals("") && finWeight != 0 ) {
			finWeight = Integer.parseInt(seed);
		}
	
		return finWeight;
	}

}
