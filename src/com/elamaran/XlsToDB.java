package com.elamaran;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class XlsToDB implements ExcelToAnyDB{
	private FileInputStream file ;
	private HSSFWorkbook workbook;
	private HSSFSheet sheet; 
	StringBuilder insert=new StringBuilder();
	StringBuilder value=new StringBuilder();
	private String sql;
	private Statement stmt;

	@Override
	public void saveToDb(Connection conn, String path, String tableName){
		List<String> columnName=new ArrayList<String>();
		boolean gotColumnName=false; 
		int column;
		
		try {		 
		 file = new FileInputStream(new File(path));
		 workbook = new HSSFWorkbook(file);
	      sheet  = workbook.getSheetAt(0);

	        //Iterate through each rows one by one
	        Iterator<Row> rowIterator = sheet.iterator();
	        while (rowIterator.hasNext())
	        {
	            Row row = rowIterator.next();
	            //For each row, iterate through all the columns
	            Iterator<Cell> cellIterator = row.cellIterator();
	            column=-1;
	            value.append("VALUES (");
	            while (cellIterator.hasNext()) 
	            {
	                Cell cell = cellIterator.next();
	                //Check the cell type and format accordingly
	                column=column+1;
	                if(column>0) {
	                	value.append(",");
	                }
	                switch (cell.getCellType()) 
	                {
	                    case NUMERIC:{
	                        //System.out.print(cell.getNumericCellValue() + "\t");
	                        if(gotColumnName==false) {
	                        	String str=Double.toString(cell.getNumericCellValue());
	                        	columnName.add(str);
	                        }else {
	                        	value.append((int)cell.getNumericCellValue());
	                        	//System.out.println(columnName.get(column)+" "+cell.getNumericCellValue());
	                        }
	                        break;
	                        }
	                    
	                    case STRING:{
	                       //System.out.print(cell.getStringCellValue() + "\t");
	                        if(gotColumnName==false) {		                        	
	                        	columnName.add(cell.getStringCellValue());
	                        }else {
	                        	value.append("'");
	                        	value.append(cell.getStringCellValue());
	                        	value.append("'");
	                        	//System.out.println(columnName.get(column)+" "+cell.getStringCellValue());
	                        }
	                        break;
	                        }
					default:
						break;
	                }
	            }
	            value.append(");");
	            if(gotColumnName==false) {
	            	insert.append("INSERT INTO ");
	            	insert.append(tableName);
	            	insert.append("(");
	            	for(int i=0;i<columnName.size();i++) {
	            		insert.append(columnName.get(i));
	            		if(i!=(columnName.size()-1)) {
	            			insert.append(",");
	            		}
	            	}
	            	insert.append(") ");
		            gotColumnName=true;
		            value=new StringBuilder();
	            }else {
	            
	            sql=insert.toString()+value.toString();
	            System.out.println(sql);
	            stmt=conn.createStatement();
	            stmt.execute(sql);
	            System.out.println("");
	            value=new StringBuilder();
	            }
	        }
	        
	        file.close();
	        workbook.close();
	    } catch (Exception e) {
	        e.printStackTrace();
	        
	    }	
	}

}
