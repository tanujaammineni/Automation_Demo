    /* 
    	Author: Artha Data Solutions Testing Team
    	Date: 2017-10-04
    	Description: This code Will Automate 13 Test cases  for Customer, Audit, Payment & Partners
    */

//Package Statement
package exe;

// Import statement for Files, Excel, Mysql & TestNG
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import junit.framework.Assert;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.AfterClass;

//Class Declaration 
public class Arthadb_Automation 
{
	public static String TC001Result,TC001Rec,TC002Result,TC002Rec,TC003Result,TC003Rec,TC004Result,TC004Rec,TC005Result,TC005Rec,TC006Result,TC006Rec,TC007Result,TC007Rec,TC008Result,TC008Rec,TC009Result,TC009Rec,TC010Result,TC010Rec,TC011Result,TC011Rec,TC012Result,TC012Rec;
	public static int TC001c,TC001RecCount,TC002c,TC002RecCount,TC003c,TC003RecCount,TC004c,TC004RecCount,TC005c,TC005RecCount,TC006c,TC006RecCount,TC007c,TC007RecCount,TC008c,TC008RecCount,TC009c,TC009RecCount,TC010c,TC010RecCount,TC011c,TC011RecCount,TC012c,TC012RecCount;
	Connection con;
	Statement stmt;
	public static int Noc;
	
	//Database Connection Statement
	@BeforeMethod
	public void DB_Details() throws ClassNotFoundException, SQLException 
	  {
		Class.forName("com.mysql.jdbc.Driver");
		  con=DriverManager.getConnection("jdbc:mysql://localhost:3306/arthadb","root","mysql"); 
		  stmt=con.createStatement(); 
	  }
	
	//ProdCodes_Check Test Script
	@Test(priority=1)
	public void ArthaDB_ProdCodes_Check_TC001() throws SQLException, RowsExceededException, WriteException, IOException 
	  {
		 StringBuffer TC001r=new StringBuffer();
		  try
		  {
			  ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.audit where ProdCodes is null"); 
			  List<String> TC001rs=new ArrayList<String>();
			  while (rs.next()) 
			  {
				  TC001c=rs.getRow();
				  TC001rs.add(rs.getString("UniqueID"));
			  }
			  if(TC001c==0)
				  {
				  	TC001Rec="AUDIT-Product Codes check is PASSED";
			        TC001RecCount=TC001rs.size();
			        TC001Result="PASS";
			        Assert.assertEquals("AUDIT-Product Codes check is PASSED", 0, TC001c);
				  }
				  else
				  {
						for(int i=0;i<TC001rs.size();i++)
						{
							  TC001r.append(TC001rs.get(i)+","); 
						}
						TC001Rec="AUDIT-Product Codes check is Failed at UniqueID="+TC001r;
				        TC001RecCount=TC001rs.size();
				        TC001Result="FAIL";
				        Assert.assertEquals("AUDIT-Product Codes check is Failed at UniqueID="+TC001r, 0, TC001c);
				  }  
		  }
		  catch (Exception e1) 
		  {
				  e1.printStackTrace();
		  } 
	}
	
	//AuditDates_Check Test Script
	@Test(priority=2)
	public void ArthaDB_AuditDates_Check_TC002() throws SQLException, RowsExceededException, WriteException, IOException 
	{
		 StringBuffer TC002r=new StringBuffer();
		  try
		  {
			  ResultSet rs=stmt.executeQuery("SELECT *FROM arthadb.audit where datediff(ExpDate,EffDate)<0"); 
			  List<String> TC002rs=new ArrayList<String>();
			  while (rs.next()) 
			  {
				  TC002c=rs.getRow();
				  TC002rs.add(rs.getString("UniqueID"));
				  
			  }
			  if(TC002c==0)
				  {
				  	TC002Rec="AUDIT-Exp and Eff Date check is PASSED";
			        TC002RecCount=TC002rs.size();
			        TC002Result="PASS";
			        Assert.assertEquals("AUDIT-Exp and Eff Date check is PASSED", 0, TC002c);
			        
				  }
				  else
				  {
						for(int i=0;i<TC002rs.size();i++)
						{
							  TC002r.append(TC002rs.get(i)+","); 
							  
						}
						
						TC002Rec="AUDIT-Exp and Eff Date check is Failed at UniqueID="+TC002r;
				        TC002RecCount=TC002rs.size();
				        TC002Result="FAIL";
				        Assert.assertEquals("AUDIT-Exp and Eff Date check is Failed at UniqueID="+TC002r, 0, TC002c);
				  }  
		  }
		  catch (Exception e1) 
		  {
				  e1.printStackTrace();
		  } 
	}
	
	//Account_Number_Check Test Script
	@Test(priority=3)
	public void ArthaDB_Account_Number_Check_TC003() throws SQLException, RowsExceededException, WriteException, IOException 
	{
		 StringBuffer TC003r=new StringBuffer();
		  try
		  { 
			  ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.payments where Account_Number like '0%'"); 
			  List<String> TC003rs=new ArrayList<String>();
			  while (rs.next()) 
			  {
				  TC003c=rs.getRow();
				  TC003rs.add(rs.getString("Invoice_Number"));
				  
			  }
			  if(TC003c==0)
				  {
				  	TC003Rec="PAYMENTS-Account Number check is PASSED";
			        TC003RecCount=TC003rs.size();
			        TC003Result="PASS";
			        Assert.assertEquals("PAYMENTS-Account Number check is PASSED", 0, TC003c);
			        
				  }
				  else
				  {
						for(int i=0;i<TC003rs.size();i++)
						{
							  TC003r.append(TC003rs.get(i)+","); 
							  
						}
						
						TC003Rec="PAYMENTS-Account Number check is Failed at Invoice_Number="+TC003r;
				        TC003RecCount=TC003rs.size();
				        TC003Result="FAIL";
				        Assert.assertEquals("PAYMENTS-Account Number check is Failed at Invoice_Number="+TC003r, 0, TC003c);
				  }  
		  }
		  catch (Exception e1) 
		  {
				  e1.printStackTrace();
		  } 
	}
	
	//ChargeType_Check Test Script
	@Test(priority=4)
	public void ArthaDB_ChargeType_Check_TC004() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  	StringBuffer TC004r=new StringBuffer();
				  try
				  {  
					  ResultSet rs=stmt.executeQuery("SELECT * fROM arthadb.payments where ChargeType not in('J','N')"); 
					  List<String> TC004rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC004c=rs.getRow();
						  TC004rs.add(rs.getString("Account_Number"));
						  
					  }
					  if(TC004c==0)
						  {
						  	TC004Rec="PAYMENTS-Charge Type Codes check is PASSED";
					        TC004RecCount=TC004rs.size();
					        TC004Result="PASS";
					        Assert.assertEquals("PAYMENTS-Charge Type Codes check is PASSED", 0, TC004c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC004rs.size();i++)
								{
									  TC004r.append(TC004rs.get(i)+","); 
									  
								}
								
								TC004Rec="PAYMENTS-Charge Type check is Failed at Account_Number(s)="+TC004r;
						        TC004RecCount=TC004rs.size();
						        TC004Result="FAIL";
						        Assert.assertEquals("PAYMENTS-Charge Type check is Failed at Account_Number(s)="+TC004r, 0, TC004c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}

	//Invoice_Number_Check Test Script
	@Test(priority=5)
	public void ArthaDB_Invoice_Number_Check_TC005() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC005r=new StringBuffer();
				  try
				  {
					  ResultSet rs=stmt.executeQuery("SELECT *  FROM arthadb.payments where Invoice_Number like '0%'"); 
					  List<String> TC005rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC005c=rs.getRow();
						  TC005rs.add(rs.getString("Account_Number"));
						  
					  }
					  if(TC005c==0)
						  {
						  	TC005Rec="PAYMENTS-Invoice Number check is PASSED";
					        TC005RecCount=TC005rs.size();
					        TC005Result="PASS";
					        Assert.assertEquals("PAYMENTS-Invoice Number check is PASSED", 0, TC005c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC005rs.size();i++)
								{
									  TC005r.append(TC005rs.get(i)+","); 
									  
								}
								
								TC005Rec="PAYMENTS-Invoice Number check is Failed at Account_Number(s)="+TC005r;
						        TC005RecCount=TC005rs.size();
						        TC005Result="FAIL";
						        Assert.assertEquals("PAYMENTS-Invoice Number check is Failed at Account_Number(s)="+TC005r, 0, TC005c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}
	
	//Transaction_Amount_Check Test Script
	@Test(priority=6)
	public void ArthaDB_Transaction_Amount_Check_TC006() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC006r=new StringBuffer();
				  try
				  {
					  ResultSet rs=stmt.executeQuery("SELECT * from arthadb.payments where Transaction_Amount=0"); 
					  List<String> TC006rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC006c=rs.getRow();
						  TC006rs.add(rs.getString("Account_Number"));
						  
					  }
					  if(TC006c==0)
						  {
						  	TC006Rec="PAYMENTS-Transaction Amount check is PASSED";
					        TC006RecCount=TC006rs.size();
					        TC006Result="PASS";
					        Assert.assertEquals("PAYMENTS-Transaction Amount check is PASSED", 0, TC006c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC006rs.size();i++)
								{
									  TC006r.append(TC006rs.get(i)+","); 
									  
								}
								
								TC006Rec="PAYMENTS-Transaction Amount check is Failed at Account_Number="+TC006r;
						        TC006RecCount=TC006rs.size();
						        TC006Result="FAIL";
						        Assert.assertEquals("PAYMENTS-Transaction Amount check is Failed at Account_Number="+TC006r, 0, TC006c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}
	
	//ChargeSign_Check Test Script	
	@Test(priority=7)
	public void ArthaDB_ChargeSign_Check_TC007() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC007r=new StringBuffer();
				  try
				  {
					  ResultSet rs=stmt.executeQuery("SELECT * from arthadb.payments where ChargeType='N' and ChargeSign!='D' and ChargeSign in('C','D')"); 
					  List<String> TC007rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC007c=rs.getRow();
						  TC007rs.add(rs.getString("Account_Number"));
						  
					  }
					  if(TC007c==0)
						  {
						  	TC007Rec="PAYMENTS-Charge Sign Codes check is PASSED";
					        TC007RecCount=TC007rs.size();
					        TC007Result="PASS";
					        Assert.assertEquals("PAYMENTS-Charge Sign Codes check is PASSED", 0, TC007c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC007rs.size();i++)
								{
									  TC007r.append(TC007rs.get(i)+","); 
									  
								}
								
								TC007Rec="PAYMENTS-Charge Sign Codes check is Failed at Account_Number="+TC007r;
						        TC007RecCount=TC007rs.size();
						        TC007Result="FAIL";
						        Assert.assertEquals("PAYMENTS-Charge Sign Codes check is Failed at Account_Number="+TC007r, 0, TC007c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}
	
	//Record_Type_Check Test Script
	@Test(priority=8)
	public void ArthaDB_Record_Type_Check_TC008() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC008r=new StringBuffer();
				  try
				  { 
					  ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.payments where (ChargeType='N' and ChargeSign='D' and Record_Type!='Debit_R5') or (ChargeType='J' and ChargeSign='D' and Record_Type!='Debit_R6') or (ChargeType='J' and ChargeSign='C' and Record_Type!='Credit_R5')"); 
					  List<String> TC008rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC008c=rs.getRow();
						  TC008rs.add(rs.getString("Invoice_Number"));
						  
					  }
					  if(TC008c==0)
						  {
						  	TC008Rec="PAYMENTS-Record Type check is PASSED";
					        TC008RecCount=TC008rs.size();
					        TC008Result="PASS";
					        Assert.assertEquals("PAYMENTS-Record Type check is PASSED", 0, TC008c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC008rs.size();i++)
								{
									  TC008r.append(TC008rs.get(i)+","); 
									  
								}
								
								TC008Rec="PAYMENTS-Record Type Check is Failed at Invoice_Number="+TC008r;
						        TC008RecCount=TC008rs.size();
						        TC008Result="FAIL";
						        Assert.assertEquals("PAYMENTS-Record Type check is Failed at Invoice_Number="+TC008r, 0, TC008c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}

	//Customer_FirstName_Check Test Script
	@Test(priority=9)
	public void ArthaDB_Customer_FirstName_Check_TC009() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC009r=new StringBuffer();
				  try
				  { 
					  ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.customers where FirstName is null"); 
					  List<String> TC009rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC009c=rs.getRow();
						  TC009rs.add(rs.getString("SSN"));
						  
					  }
					  if(TC009c==0)
						  {
						  	TC009Rec="CUSTOMERS-First Name check is PASSED";
					        TC009Result="PASS";
					        Assert.assertEquals("CUSTOMERS-First Name check is PASSED", 0, TC009c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC009rs.size();i++)
								{
									  TC009r.append(TC009rs.get(i)+","); 
									  
								}
								
								TC009Rec="CUSTOMERS-First Name Check is Failed at SSN="+TC009r;
						        TC009RecCount=TC009rs.size();
						        TC009Result="FAIL";
						        Assert.assertEquals("CUSTOMERS-First Name check is Failed at SSN="+TC009r, 0, TC009c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}
	
	//Customer_SSN_Check Test Script
	@Test(priority=10)
	public void ArthaDB_Customers_SSN_Check_TC010() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC010r=new StringBuffer();
				  try
				  { 
					  ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.customers where SSN is null or length(SSN)!=11"); 
					  List<String> TC010rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC010c=rs.getRow();
						  TC010rs.add(rs.getString("CustomerPK"));
						  
					  }
					  if(TC010c==0)
						  {
						  	TC010Rec="CUSTOMERS- SSN check is PASSED";
					        TC010Result="PASS";
					        Assert.assertEquals("CUSTOMERS- SSN check is PASSED", 0, TC010c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC010rs.size();i++)
								{
									  TC010r.append(TC010rs.get(i)+","); 
									  
								}
								
								TC010Rec="CUSTOMERS- SSN Check is Failed at CustomerPK="+TC010r;
						        TC010RecCount=TC010rs.size();
						        TC010Result="FAIL";
						        Assert.assertEquals("CUSTOMERS- SSN check is Failed at CustomerPK="+TC010r, 0, TC010c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}
	
	//Customer_Record_Type_Check Test Script
	@Test(priority=11)
	public void ArthaDB_Record_Type_Check_TC011() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC011r=new StringBuffer();
				  try
				  {
					  ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.partners where Record_Type is null"); 
					  List<String> TC011rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC011c=rs.getRow();
						  TC011rs.add(rs.getString("Partner_Account_Number"));
						  
					  }
					  if(TC011c==0)
						  {
						  	TC011Rec="PARTNERS-Record Type check is PASSED";
					        TC011Result="PASS";
					        Assert.assertEquals("PARTNERS-Record Type check is PASSED", 0, TC011c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC011rs.size();i++)
								{
									  TC011r.append(TC011rs.get(i)+","); 
									  
								}
								
								TC011Rec="PARTNERS-Record Type Check is Failed at Partner_Account_Number="+TC011r;
						        TC011RecCount=TC011rs.size();
						        TC011Result="FAIL";
						        Assert.assertEquals("PARTNERS-Record Type check is Failed at Partner_Account_Number="+TC011r, 0, TC011c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	}
	
	//Customer_Premise_City_Check Test Script
	@Test(priority=12)
	public void ArthaDB_Premise_City_Check_TC012() throws SQLException, RowsExceededException, WriteException, IOException 
	{
			  StringBuffer TC012r=new StringBuffer();
				  try
				  {
					  ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.partners where (Premise_City is null or Premise_City=' ' or Premise_City='')"); 
					  List<String> TC012rs=new ArrayList<String>();
					  while (rs.next()) 
					  {
						  TC012c=rs.getRow();
						  TC012rs.add(rs.getString("Partner_Account_Number"));
						  
					  }
					  if(TC012c==0)
						  {
						  	TC012Rec="PARTNERS-City check is PASSED";
					        TC012Result="PASS";
					        Assert.assertEquals("PARTNERS-City check is PASSED", 0, TC012c);
					        
						  }
						  else
						  {
								for(int i=0;i<TC012rs.size();i++)
								{
									  TC012r.append(TC012rs.get(i)+","); 
									  
								}
								
								TC012Rec="PARTNERS-City Check is Failed at Partner_Account_Number="+TC012r;
						        TC012RecCount=TC012rs.size();
						        TC012Result="FAIL";
						        Assert.assertEquals("PARTNERS-City check is Failed at Partner_Account_Number="+TC012r, 0, TC012c);
						  }  
				  }
				  catch (Exception e1) 
				  {
						  e1.printStackTrace();
				  } 
	} 
	
	//Customer_PrimaryPhoneNumber_Check Test Script
	@Test(priority=13)
	public void ArthaDB_PrimaryPhoneNumber_TC013() throws SQLException
	{
		ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.customers where PrimaryPhoneNumber is null");
		List<String> Phone_rs=new ArrayList<String>();
		while(rs.next())
		{
			Noc=rs.getRow();
			Phone_rs.add(rs.getString("SSN"));
		}
		if(Noc==0)
		{
			Assert.assertEquals("Customer - PrimaryPhoneNumber check is Passed",0,Noc);
		}
		else
		{
			StringBuffer Phone_r=new StringBuffer();
			for(int k=0;k<Phone_rs.size();k++)
			{
				Phone_r.append(Phone_rs.get(k)+",");
			}
			Assert.assertEquals("Customer - PrimaryPhoneNumber check is Failed at SSN="+Phone_r, 0, Noc);
		}
	}
}
