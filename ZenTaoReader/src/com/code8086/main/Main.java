package com.code8086.main;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.*;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;



public class Main
{
	public static void main (String[] args) throws Exception
	{
		Connection conn = null;
        String sql;
        String url = "jdbc:mysql://localhost:3306/zentao?"
                + "user=root&password=&useUnicode=true&characterEncoding=UTF8&zeroDateTimeBehavior=convertToNull";
        Class.forName("com.mysql.jdbc.Driver");
        System.out.println("驱动加载成功!");
        conn = DriverManager.getConnection(url);
        Statement stmt = conn.createStatement();
        sql = "select * from zt_task";
        //int result = stmt.executeUpdate(sql);
        ResultSet rs = stmt.executeQuery(sql);
        
        String file_path = "output.xls";
        // 创建Excel工作薄   
        WritableWorkbook wwb;   
         // 新建立一个jxl文件,即在d盘下生成testJXL.xls   
        OutputStream os = new FileOutputStream(file_path);   
        wwb=Workbook.createWorkbook(os);    
        // 添加第一个工作表并设置第一个Sheet的名字   
        WritableSheet sheet = wwb.createSheet("任务详情", 0);
        
        int index = 1;
        
        while (rs.next())
        {
        	//System.out.println("查询成功!");
        	/*System.out.print(rs.getString("id") + " ");
        	System.out.print(rs.getString("project") + " ");
        	System.out.print(rs.getString("module") + " ");
        	System.out.print(rs.getString("story") + " ");
        	System.out.print(rs.getString("storyVersion") + " ");
        	System.out.print(rs.getString("fromBug") + " ");
        	System.out.print(rs.getString("name") + " ");
        	System.out.print(rs.getString("type") + " ");
        	System.out.print(rs.getString("pri") + " ");
        	System.out.print(rs.getString("estimate") + " ");
        	System.out.println();*/
        	//Label label = new Label(index, 0, rs.getString("id"));
        	for (int i = 1; i <= 32; i++)
        	{
        		Label label = new Label(i, index, rs.getString(i));
        		sheet.addCell(label);
        	}
        	index++;
        }
        System.out.println("Done!");
        wwb.write();
        wwb.close();
	}
}
