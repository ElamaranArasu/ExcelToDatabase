package com.elamaran;

import java.sql.Connection;

public interface ExcelToAnyDB {
	public void saveToDb(Connection conn,String filepath,String tableName);
}
