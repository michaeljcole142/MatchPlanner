package statsProcessor;

import java.util.*;



/*
 *  This is a class to hold dual meets.
 *  Notice that it extends Event.  An event is 
 *  a class that holds the things that are common to both a
 *  Dual Meet and a Tournament.
 */
class GSDualMeet extends DualMeet {
	
	private static String[] GSWeightList = { "70","75","80","85","90","95","100","106","112","118","124","130","140","150","175","240" };
	private static int[] GSWeightListInt = { 70,75,80,85,90,95,100,106,112,118,124,130,140,150,175,240,9999 };
	
	GSDualMeet() {
		this.resetWeights();
		this.initialize(GSWeightList);
	}
	public static String[] getWeightList() { return GSDualMeet.GSWeightList; }
	public static int[] getWeightListInt() { return GSDualMeet.GSWeightListInt; }
	
}
