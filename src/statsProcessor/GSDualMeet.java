package statsProcessor;

import java.util.*;



/*
 *  This is a class to hold dual meets.
 *  Notice that it extends Event.  An event is 
 *  a class that holds the things that are common to both a
 *  Dual Meet and a Tournament.
 */
class GSDualMeet extends DualMeet {
	
//	private static String[] GSWeightList = { "70","75","80","85","90","95","100","106","112","118","124","130","140","150","175","240" };
//	private static int[] GSWeightListInt = { 70,75,80,85,90,95,100,106,112,118,124,130,140,150,175,240,9999 };
	
	//Central Jersey Weights
//	private static String[] GSWeightList = { "50","53","57","60","63","67","70","73","77","80","85","90","95","102","112","125","285" };
//	private static int[] GSWeightListInt = { 50,53,57,60,63,67,70,73,77,80,85,90,95,102,112,125,285,9999 };
	
	//NW Jersey Weights
	private static String[] GSWeightList = { "46","49","52","55","58","61","64","67","70","73","76","80","85","90","95","102","115","135","165","200" };
	private static int[] GSWeightListInt = { 46,49,52,55,58,61,64,67,70,73,76,80,85,90,95,102,115,135,165,200,9999 };
	
	
	GSDualMeet() {
		this.resetWeights();
		this.initialize(GSWeightList);
	}
	public static String[] getWeightList() { return GSDualMeet.GSWeightList; }
	public static int[] getWeightListInt() { return GSDualMeet.GSWeightListInt; }
	
}
