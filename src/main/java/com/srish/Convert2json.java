package com.srish;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

public class Convert2json {
	public static void main(String[] args) {
		
		
    List<Customer> customers = readExcelFile("customers-1.xlsx");
	    
	    String input_param = "Jack Smith";
	    
	    String jsonString = convertObjects2JsonString(input_param,customers);
	    
	    System.out.println("here-->"+ jsonString);
	  }
	  

	  private static List<Customer> readExcelFile(String filePath){
	    try {
	      FileInputStream excelFile = new FileInputStream(new File(filePath));
	        Workbook workbook = new XSSFWorkbook(excelFile);
	     
	        Sheet sheet = workbook.getSheet("Customers");
	        Iterator<Row> rows = sheet.iterator();
	        
	        List<Customer> lstCustomers = new ArrayList<Customer>();
	        
	        int rowNumber = 0;
	        while (rows.hasNext()) {
	          Row currentRow = rows.next();
	          
	        
	          if(rowNumber == 0) {
	            rowNumber++;
	            continue;
	          }
	          
	          Iterator<Cell> cellsInRow = currentRow.iterator();
	 
	         Customer cust= new Customer();
	          
	          int cellIndex = 0;
	          while (cellsInRow.hasNext()) {
	            Cell currentCell = cellsInRow.next();
	            
	            if(cellIndex==0) { 
	              cust.setId((int) currentCell.getNumericCellValue());
	            } else if(cellIndex==1) { 
	              cust.setName(currentCell.getStringCellValue());
	            } else if(cellIndex==2) { 
	              cust.setAddress(currentCell.getStringCellValue());
	            } else if(cellIndex==3) {
	              cust.setAge((int) currentCell.getNumericCellValue());
	            }
	            
	            cellIndex++;
	          }
	          
	          lstCustomers.add(cust);
	        }
	        
	        
	        // workbook.close();
	        
	        return lstCustomers;
	        } catch (IOException e) {
	          throw new RuntimeException("FAIL! -> message = " + e.getMessage());
	        }
	  }
	  
	  private static String convertObjects2JsonString(String input_param, List<Customer> customers) {
	      ObjectMapper mapper = new ObjectMapper();
	      String jsonString = "" ;

	      try {
	    	  //comparing Customer name, if you want any other filed to compare
	    	  // get that filed in place of ee.getName()
		      List<Customer> list = customers.stream().filter(
		    		  ee-> ee.getName().equals(input_param)).collect(Collectors.toList());
	    	  if(list.size()==0) {
	    		  throw new IllegalAccessError();
	    	  }
	        jsonString = mapper.writeValueAsString(list);
	      } catch (JsonProcessingException e) {
	        e.printStackTrace();
	      }
	      
	      return jsonString; 
	}

}
