package src;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream; 
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData; 
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
//import java.text.DateFormat;
//import java.text.SimpleDateFormat;
import java.util.ArrayList;
//import java.util.Calendar;
//import java.util.Date;
import java.util.Properties; 
import java.util.StringTokenizer;



//import javax.activation.*;

import org.apache.poi.hssf.usermodel.HSSFCell; 
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat; 
import org.apache.poi.hssf.usermodel.HSSFFont; 
import org.apache.poi.hssf.usermodel.HSSFRow; 
import org.apache.poi.hssf.usermodel.HSSFSheet; 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;;
public class DashboardUtil {
	
	FileOutputStream out = null;
	HSSFWorkbook wb = null;
	HSSFSheet ws = null;
	HSSFRow r = null;
	HSSFCell c = null;
	HSSFCellStyle cs1 = null;
	HSSFCellStyle cs2 = null; 
	HSSFCellStyle cs3 = null;
	HSSFDataFormat df = null;
	HSSFFont f1 = null;
	HSSFFont f2 = null;
	
	Connection con;
	Statement st;
	ResultSet rs;
	
	static Properties pr;
	static String filename = "";
	public void createExcelFile(String filename) throws Exception
	{
		out = new FileOutputStream(filename);
		wb = new HSSFWorkbook();
		ws = this.wb.createSheet();
		cs1 = this.wb.createCellStyle();
		cs2 = this.wb.createCellStyle();
		cs3 = this.wb.createCellStyle();
		df = this.wb.createDataFormat();
		f1 = this.wb.createFont();
		f2 = this.wb.createFont();
		
		f1.setFontHeightInPoints((short)10);
		f1.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		
		f2.setFontHeightInPoints((short)13);
		f2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		f2.setColor(HSSFFont.COLOR_RED);
		
		cs1.setFont(f1);
		cs1.setDataFormat(df.getFormat("text"));
		
		cs2.setFont(f2);
		cs2.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));
		
		cs3.setFont(f1);
		cs3.setDataFormat(df.getFormat("#,##0.0"));
		
		wb.setSheetName(0, "QryDetail", HSSFWorkbook.ENCODING_UTF_16);
		
	}
	
	public void populateData(ArrayList<String> ColumnArraylist, ResultSet rset,String filename)
	{
		try
		{
			r = ws.createRow(0);
			for(int i = 0;i<ColumnArraylist.size();i++)
			{
				c = r.createCell((short)i);
				c.setCellStyle(cs2);
				c.setEncoding(HSSFCell.ENCODING_UTF_16);
				c.setCellValue(ColumnArraylist.get(i).toString());
			}
			//c.setCellValue("its me");
			ResultSetMetaData rsetmeta = rset.getMetaData();
			int count=1;
			while(rset.next())
			{
				r = ws.createRow(count);
				for(int i = 0;i<ColumnArraylist.size();i++)
				{
					c = r.createCell((short)i);
					if(rsetmeta.getColumnType(i+1)==Types.INTEGER)
					{
					c.setCellStyle(cs1);
					c.setEncoding(HSSFCell.ENCODING_UTF_16);
					System.out.println("************Check"+rset.getInt(i+1));
					c.setCellValue(rset.getInt(i+1));
					}
					else
					{
						c.setCellStyle(cs1);
						c.setEncoding(HSSFCell.ENCODING_UTF_16);
						c.setCellValue(rset.getString(i+1));
					}
				}
				count++;
			}
			wb.write(out);
			out.close();
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
	}
	
	public Statement createConnection()
	{
		try
		{
		Class.forName("oracle.jdbc.driver.OracleDriver");
		con = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:xe", "hr","sajal");
		st = con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_READ_ONLY);
		System.out.println("Created DB Connection....");
		}
		catch (ClassNotFoundException e) 
		{
		
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (SQLException e) 
		{
		// TODO Auto-generated catch block
		e.printStackTrace();
		}
		return st;
	}
	
	public ResultSet getdata(Statement st)
	{
		String sql;
	      sql = "select employees_id,first_name,last_name from employee order by employees_id";
	      try {
			rs = st.executeQuery(sql);
		      /*while(rs.next()){
	          //Retrieve by column name
	          int id  = rs.getInt("employees_id");
	          String fname = rs.getString("first_name");
	          String last = rs.getString("last_name");

	          //Display values
	          System.out.print("ID: " + id);
	          System.out.print(", First: " + fname);
	          System.out.println(", Last: " + last);
	          
	      }*/
	      } catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	      return rs;
	}
	
	public void genrateReport(ResultSet rs,String filename,String prLoc)
	{
		File file = new File(prLoc);
		FileInputStream fileInput;
		try {
			fileInput = new FileInputStream(file);
			pr = new Properties();
			pr.load(fileInput);
			System.out.println("Property Loaded Successfully....");
			
		}
	catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
				
		String columns = pr.getProperty("column");
		System.out.println(columns);
		StringTokenizer columnlist = new StringTokenizer(columns,",");
		ArrayList<String> arl = new ArrayList<String>();
		while(columnlist.hasMoreElements())
		{
			arl.add(columnlist.nextToken());
		}
		
		File f = new File("D:\\Report\\");
		f.mkdir();
		try {
			filename=f.getAbsolutePath()+"\\" + filename;
			createExcelFile(filename);
			populateData(arl,rs,filename);
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	public void closeAll()
	{
		try {
			getdata(createConnection()).close();
			createConnection().close();
			System.out.println("Connection Closed Successfully......");
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void main(String s[]) throws Exception
	{
		
		DashboardUtil gw = new DashboardUtil();
		//gw.createExcelFile("D:\\Sajal\\MyTest.xls");
		//gw.populateData();
		//gw.createConnection();
		gw.genrateReport(gw.getdata(gw.createConnection()),"MyTest.xls","D:\\Sajal\\Workspace\\packageTest\\src\\packageTest\\IW.properties");
		gw.closeAll();
		
	}
}

