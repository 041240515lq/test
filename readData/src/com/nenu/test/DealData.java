package com.nenu.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.ArrayList;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

/**
 * Servlet implementation class DealData
 */
@WebServlet("/DealData")
public class DealData extends HttpServlet {
	private static final long serialVersionUID = 1L;
    

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		ArrayList<String> strList;
		 try {
			 strList = new ArrayList<String>();
			 response.setContentType("text/html;charSet=utf-8"); 
			 PrintWriter out = response.getWriter();   
			 String tel=request.getParameter("pn");
			 File file = new File("C:\\Users\\liqua\\Desktop\\test.xls");  
			    InputStream in = new FileInputStream(file);  
			    Workbook workbook = Workbook.getWorkbook(in);  
			    //��ȡ��һ��Sheet��  
			    Sheet rs = workbook.getSheet(0);  
			      
			    //���Ǽȿ���ͨ��Sheet����������������Ҳ����ͨ���±��������������ͨ���±������ʵĻ���Ҫע���һ�����±��0��ʼ����������һ����  
			    //��ȡ��һ�У���һ�е�ֵ   
			 /*   Cell c00 = rs.getCell(0, 0);   
			    String strc00 = c00.getContents();   
			    //��ȡ��һ�У��ڶ��е�ֵ   
			    Cell c10 = rs.getCell(1, 0);   
			    String strc10 = c10.getContents(); 
			    System.out.println(strc00+"        "+strc10);*/
			    
			    Cell[] cells = rs.getColumn(3);  
			    for(int i=1;i<cells.length;i++){
			    	String temp = cells[i].getContents();
			    	if(temp.startsWith("'")){
			    		strList.add(temp.substring(1)); 
			    	}else{
			    		strList.add(temp); 
			    	}
			    	
			    }
			    int num=-1;
			    for(int j=0;j<strList.size();j++){
			    	if(tel.equals(strList.get(j))){
			    		num=j;
			    		break;
			    	}
			    }
			    Cell[] orderCells = rs.getColumn(4);  
			    out.write(orderCells[num+1].getContents()+"\r\n");
			    
			 System.out.println(tel);
			  
			} catch (Exception e) {
			// TODO: handle exception
		}  
		 }
}
