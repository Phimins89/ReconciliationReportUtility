package com.kbc.efs.util;

public class CSVentry implements Comparable<CSVentry>{

		public int batchNum, numOfDocuments;
		public String letterType, result;

		public CSVentry(int batchNum, String letterType, int numOfDocuments) {
		    super();
		    this.batchNum = batchNum;
		    this.letterType = letterType;
		    this.numOfDocuments = numOfDocuments;
		  	// System.out.println(this.batchNum);
		}

		public int getBatchNumber() {
		    return batchNum;
		}

		public String getLetterType() {
		    return letterType;
		}

		public int getNumOfDocuments() {
		    return numOfDocuments;
		}

		public String getComparisonResult() {
		    return result;
		}

		@Override
		public int compareTo (CSVentry s) {

		   return ((Integer)this.batchNum).compareTo(s.batchNum);
		   
		}

		
	

		}

