/**
 * 
 */
package statsProcessor;

/**
 * @author Mike
 * 
 * This file contains a wrestling language class.  It holds all the terms 
 * that we will find across classes.
 */
class WrestlingLanguage {
 
  
   public static enum MatchResultType { FALL, TECH, MD, DEC, FFT, INJ, DQ, NC, MFFT };
   //public static enum MatchResultType { FALL, MD, DEC, FFT, INJ, TECH };
   public static enum WinOrLose { WIN, LOSS };
   public static enum Grade { G5, G6, G7, G8, FR, SO, JR, SR };
   public static enum Gender { M, F };

   public static enum Prestige { 
								Dnever, 
								DPrelim,
								DQuarters,
								D4th,
								R2v3,
								RWB1,
								R6th,
								R5th,
								SWB1,
								SWB2,
								SWB3,
								S8th,
								S7th,
								S6th,
								S5th,
								S4th,
								S3rd,
								S2nd,
								S1st
										};

  }