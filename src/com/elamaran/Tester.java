package com.elamaran;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

//import com.elamaran.ExcelToAnyDB;
//import com.elamaran.XlsToDB;


public class Tester {

	public static void main(String[] args) throws SQLException {
		 ExcelToAnyDB xl2db=new XlsToDB();
		 Connection conn=null;
		 try{
			 conn= DriverManager.getConnection("jdbc:postgresql://localhost:5432/testdb2", "ela", "root");
			 System.out.println("Database connected");
			 xl2db.saveToDb(conn, "/home/ela/java/other files/test.xls", "testtable");			 
		 }catch(Exception e) {
			 e.printStackTrace();
	         System.err.println(e.getClass().getName()+": "+e.getMessage());
	         System.exit(0);
		 }finally {
			 if(conn!=null) {
				 conn.close();
			 }
		 }
	}
}
