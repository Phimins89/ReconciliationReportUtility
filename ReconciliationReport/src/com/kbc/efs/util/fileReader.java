package com.kbc.efs.util;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class fileReader {
	
	 private static final DateFormat sdf = new SimpleDateFormat("ddMMyyyy");

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	
//		final File MI_directory = new File("P:\\ReconciliationReport\\MI_Report\\");
//		final File RRD_directory = new File("P:\\ReconciliationReport\\RRD_Report\\");
		
		final File MI_directory = new File(System.getenv("RECRPT_DIR") + "MI_Report\\");
		final File RRD_directory = new File(System.getenv("RECRPT_DIR") + "RRD_Report\\");
		final File MI_directory_Archive = new File(System.getenv("RECRPT_DIR") + "Archived\\MI_Report\\");
		final File RRD_directory_Archive = new File(System.getenv("RECRPT_DIR") + "Archived\\RRD_Report\\");
		
		File[] listOfFiles_MI = MI_directory.listFiles();
		File[] listOfFiles_LN = RRD_directory.listFiles();
		String File1 = "";
		String File2 = "";
		String WarningText = "Reconciliation Successful. No discrepancies found";
	    String BatchesFailed = null;
		
		int index = 0;
		 for (File f : listOfFiles_MI) {
	//	listOfFiles_MI.forEach(f->System.out.println(f));
		
		      if (f.isFile()) {
		    	 try{  index++;
		    	  File1 +=  new String(Files.readAllBytes(Paths.get(f.getPath())));
	    	  }
	    	  catch (IOException e) {
	    	      System.out.println(e);
	    	    }	 
		 
		      }
		 }
		 index = 0;
		 for (File f : listOfFiles_LN) {
		      if (f.isFile()) {
		    	  try {
		    	  index++;
		    	  
		    	  File2 +=  new String(Files.readAllBytes(Paths.get(f.getPath())));
		    	  }
		    	  catch (IOException e) {
		    	      System.out.println(e);
		    	    }
		      }
		//      System.out.println(File2);
		 }

		 ArrayList<CSVentry> FileOne = FileComparison(File2);
		 ArrayList<CSVentry> FileTwo = FileComparison(File1);
		 
		int issueCounter = 0; 
	    for(int i=0;i<FileOne.size();i++){
	    	
	       if (FileOne.get(i).batchNum != FileTwo.get(i).batchNum){
	    	   WarningText = "The following batches were sent by EFS but were not present in the MI Report :";
	    	   BatchesFailed += FileOne.get(i).batchNum + " , Letter Type => " + FileOne.get(i).letterType + " , Number of Docs => " + FileOne.get(i).numOfDocuments + "\n";
	    	   BatchesFailed = BatchesFailed.replace("null", "");
	    	   issueCounter++;
	    	   FileOne.remove(i);
	       }
	       if(FileOne.get(i).numOfDocuments != FileTwo.get(i).numOfDocuments && FileOne.get(i).letterType != FileTwo.get(i).letterType){
			   FileOne.get(i).result = "DIFFERENT";
			   WarningText = "Results are not identical. Please check below";
			   issueCounter++;
		   }
		   else{ 
			   FileOne.get(i).result = "IDENTICAL";
		   }
	   }  
	    if(issueCounter < 1 ){
	    	WarningText = "Reconciliation Successful. All batches received";
	    }
	 
	    //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook(); 
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("RRD vs MI Report");
          
        //This data needs to be written (Object[])
        Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();
        
        data.put(0, new Object[] {WarningText});  
        data.put(1, new Object[] {BatchesFailed}); 
        data.put(2, new Object[] {"Letter Type", "Batch Number(RRD)", "Doc Count", "Comparison Results(RRD vs MI)", "Letter Type", "Batch Number(MI)", "Doc Count"});
        
        for(int i=0;i < FileOne.size();i++){
        	data.put(3+i ,new Object[] {FileOne.get(i).letterType, FileOne.get(i).batchNum, FileOne.get(i).numOfDocuments, FileOne.get(i).result, FileTwo.get(i).letterType, FileTwo.get(i).batchNum, FileTwo.get(i).numOfDocuments});
        }
     
        //Iterate over data and write to sheet
        Set<Integer> keyset = data.keySet();
               
        int rownum = 0;
        Row row = null;
        for (Integer key : keyset)
        {
        	
            
            
            row = sheet.createRow(rownum++);
           
            
            Object [] objArr = data.get(key);
            int cellnum = 0;
            
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
                           
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
               
               sheet.autoSizeColumn(cellnum);
               
             }
           
     //       row.getCell(0).setCellStyle(style);
            
        }
        Font boldfont = workbook.createFont();
        CellStyle style = workbook.createCellStyle();
        boldfont.setFontHeightInPoints((short)12);
        boldfont.setBold(true);
        style.setFont(boldfont);
        
        Font redfont = workbook.createFont();
        CellStyle style2 = workbook.createCellStyle();
        style2.setWrapText(true);
        redfont.setColor(IndexedColors.RED.getIndex());
        style2.setFont(redfont);
        
        sheet.getRow(0).getCell(0).setCellStyle(style);
        sheet.getRow(1).getCell(0).setCellStyle(style2);
        sheet.autoSizeColumn(0);
        
        for(int i=0;i<=6;i++)
        sheet.getRow(2).getCell(i).setCellStyle(style);
       
               
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        
           // Condition 1: Cell Value is equal to green (Green Font)
           ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"IDENTICAL\"");
           FontFormatting fill1 = rule1.createFontFormatting();
           fill1.setFontColorIndex(IndexedColors.GREEN.index);
           CellRangeAddress[] regions = {CellRangeAddress.valueOf("D1:D200")};
           sheetCF.addConditionalFormatting(regions, rule1);
           
           // Condition 1: Cell Value is equal to RED (Red Font)
           ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"DIFFERENT\"");
           FontFormatting fill2 = rule2.createFontFormatting();
           fill2.setFontColorIndex(IndexedColors.RED.index);
           CellRangeAddress[] regions2 = {CellRangeAddress.valueOf("D1:D200")};
           sheetCF.addConditionalFormatting(regions2, rule2);
          
           try
           {
	        	Date date = new Date();
	    //    	FileOutputStream out = new FileOutputStream(new File("P://ReconciliationReport//Reconciled//ReconciliationReport_"+sdf.format(date)+".xlsx"));
	        	//Write the workbook in file system on server
	        	FileOutputStream out = new FileOutputStream(new File(System.getenv("RECRPT_DIR") + "Reconciled/ReconciliationReport_"+sdf.format(date)+".xlsx"));
	            workbook.write(out);
	            out.close();
	            System.out.println("ReconciliationReport_"+sdf.format(date)+".xlsx written successfully on disk.");
	        } 
	        catch (Exception e) 
	        {
	            e.printStackTrace();
	        }
        
        if(listOfFiles_MI.length > 0){
        	for (int i=0;i<listOfFiles_MI.length;i++){
		       	System.out.println("MI File to be moved from : " + listOfFiles_MI[i].toPath());
		    	System.out.println("MI File to be moved to : " + Paths.get(MI_directory_Archive.getPath(), listOfFiles_MI[i].getName()));
				Files.move(listOfFiles_MI[i].toPath(), Paths.get(MI_directory_Archive.getPath(), listOfFiles_MI[i].getName()), StandardCopyOption.REPLACE_EXISTING);	
        	}
		    for (int i=0;i<listOfFiles_LN.length;i++){
		       	System.out.println("RRD File to be moved from : " + listOfFiles_LN[i].toPath());
		        System.out.println("RRD File to be moved to : " + Paths.get(RRD_directory_Archive.getPath() , listOfFiles_LN[i].getName()));
				Files.move(listOfFiles_LN[i].toPath(), Paths.get(RRD_directory_Archive.getPath(), listOfFiles_LN[i].getName()), StandardCopyOption.REPLACE_EXISTING);
			}
        }  
        workbook.close(); 
	}

	private static ArrayList<CSVentry> FileComparison(String RecFile) {
		// TODO Auto-generated method stub
		String[] entries = null;
		String line = "";
	    String cvsSplitBy = ",";
	   
	    int letterTypePos = 0, batchNumPos = 0, mailpiecePos = 0, numOfPagesPos = 0;
	    
	   
	    ArrayList<CSVentry> listOfLists = new ArrayList<>();
	    
	    int index = 0;
	    
		try (BufferedReader br = new BufferedReader(new StringReader(RecFile))) {

            while ((line = br.readLine()) != null) {
                // use comma as separator  	
            	
            	 if(line.contains("BatchNo")){
          	    	cvsSplitBy = " ";
          	    	batchNumPos   = 2;
          	    	letterTypePos = 7;
          	    	mailpiecePos = 11;
          	    } else {
          	    	cvsSplitBy = ",";
          	    	}
            	     	
            	entries  = line.split(cvsSplitBy);
            	
            	for(int i=0; i < entries.length; i++){
         
                	 if(entries[i].contains("LetterType")){
                 		 letterTypePos = i;
                  	 }
                	 if(entries[i].contains("KBC-Batch")){
                  		 batchNumPos = i;
                  	 }
                	 if(entries[i].contains("Mailpiece-cnt")){
                  		 mailpiecePos = i;
                	 }
                	
             	} 
            	        	
            	if(entries[0].equals("BatchNo")){
            		//System.out.println("Batch Number1 : "+ entries[batchNumPos] + " | " + entries[letterTypePos] + " | " + entries[mailpiecePos]);
             		listOfLists.add(new CSVentry(Integer.valueOf(entries[batchNumPos]), entries[letterTypePos], Integer.valueOf(entries[mailpiecePos])));
            	}else {
            		Pattern p = Pattern.compile("[A-Z+]");
            		Matcher m = p.matcher(entries[batchNumPos]);
            		
            	if(index > 0 && !m.find()){
            		//System.out.println("Batch Number2 : "+ entries[batchNumPos] + " | " + entries[letterTypePos] + " | " + entries[mailpiecePos]);
            		listOfLists.add(new CSVentry(Integer.valueOf(entries[batchNumPos]), entries[letterTypePos], Integer.valueOf(entries[mailpiecePos])));
            	}
            	}
                index++;
                        	
            }
            
           Collections.sort(listOfLists);
               	
          
            
        }catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println("It Didn't Work :(");
        }
		return listOfLists;
		
	}
}

