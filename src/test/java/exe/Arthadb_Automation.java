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

import org.testng.Reporter;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.AfterClass;

//Class Declaration 
public class Arthadb_Automation
{
	//public static int Noc;
	Connection con;
	Statement stmt;

	//Database Connection Statement
	@BeforeMethod
	public void DB_Details() throws ClassNotFoundException, SQLException 
	{
		Class.forName("com.mysql.jdbc.Driver");
		con=DriverManager.getConnection("jdbc:mysql://localhost:3306/arthadb","root","mysql"); 
		stmt=con.createStatement(); 
	}

	/*
	 * Customer Name Check
	 * 			1. Here we are Validating Both First & Last Name
	 * 			2. Both are Mandatory
	 * 			3. Both Names Allow Only Alphabets (a to z / A to Z)
	 */
	@Test(priority=1)
	public void ArthaDB_Customer_FirstName_Check_TC001() throws SQLException, RowsExceededException, WriteException, IOException 
	{
		StringBuffer TC009r=new StringBuffer();
		try
		{ 
			int Noc=0; 
			ResultSet rs=stmt.executeQuery("select * from customers where FirstName regexp '^[^A-Za-z]' or FirstName is null"); 
			List<String> TC009rs=new ArrayList<String>();
			while (rs.next()) 
			{
				Noc=rs.getRow();
				TC009rs.add(rs.getString("SSN"));

			}
			if(Noc==0)
			{
				Assert.assertEquals("CUSTOMERS-FirstName check is PASSED", 0, Noc);
			}
			else
			{
				for(int i=0;i<TC009rs.size();i++)
				{
					TC009r.append(TC009rs.get(i)+","); 	  
				}
				Assert.assertEquals("CUSTOMERS-FirstName check is Failed at SSN="+TC009r, 0, Noc);
			}  
		}
		catch (Exception e1) 
		{
			e1.printStackTrace();
		} 
	}

	/*
	 * Here we will Check Customer SSN Format
	 * 		1. SSN is Mandatory
	 * 		2. SSN is in Specified Format like 123-45-6796 
	 */
	@Test(priority=2)
	public void ArthaDB_Customers_SSN_Check_TC002() throws SQLException, RowsExceededException, WriteException, IOException 
	{
		StringBuffer TC010r=new StringBuffer();
		try
		{ 
			int Noc=0; 
			ResultSet rs=stmt.executeQuery("SELECT * FROM customers WHERE SSN not RLIKE '[[:<:]][0-9]{3}[[:>:]](-)[[:<:]][0-9]{2}[[:>:]](-)[[:<:]][0-9]{4}[[:>:]]' or SSN is null"); 
			List<String> TC010rs=new ArrayList<String>();
			while (rs.next()) 
			{
				Noc=rs.getRow();
				TC010rs.add(rs.getString("CustomerPK"));

			}
			if(Noc==0)
			{
				Assert.assertEquals("CUSTOMERS- SSN check is PASSED", 0, Noc);

			}
			else
			{
				for(int i=0;i<TC010rs.size();i++)
				{
					TC010r.append(TC010rs.get(i)+","); 
				}
				Assert.assertEquals("CUSTOMERS- SSN check is Failed at CustomerPK="+TC010r, 0, Noc);
			}  
		}
		catch (Exception e1) 
		{
			e1.printStackTrace();
		} 
	}

	/*
	 * Customer Phone number is mandatory and it is in specified format (for example +1 333 333 4444)
	 * Input Excel Phone number does not contain PLUS (+) symbol, it contain hyphen(-)
	 * Need to replay hyphen symbol into Space. 
	 */
	@Test(priority=3)
	public void ArthaDB_PrimaryPhoneNumber_TC003() throws SQLException
	{
		int Noc=0; 
		ResultSet rs=stmt.executeQuery("SELECT * FROM customers WHERE PrimaryPhoneNumber not REGEXP '[+][1] [0-9]{3} [0-9]{3} [0-9]{4}' or PrimaryPhoneNumber is null");
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

	/*
	 * Customer Details Update Check
	 * Here If any Customer Change Those respective Details, we will find those SSN Numbers.
	 */
	@Test(priority=4)
	public void ArthaDB_Customer_Details_Update_Check_TC004() throws SQLException
	{
		int Noc=0; 
		ResultSet rs=stmt.executeQuery("SELECT * FROM arthadb.customers_audit_table where action='Update'");
		List<String> Update_rs=new ArrayList<String>();
		while(rs.next())
		{
			Noc=rs.getRow();
			Update_rs.add(rs.getString("SSN"));
		}
		if(Noc==0)
		{
			Assert.assertEquals("Customer - Details check is Passed",0,Noc);
		}
		else
		{
			StringBuffer Update_r=new StringBuffer();
			for(int k=0;k<Update_rs.size();k++)
			{
				Update_r.append(Update_rs.get(k)+",");
			}
			Reporter.log("Customer - Details are updated at Customer SSN= "+Update_r);
			//Assert.assertEquals("Customer - Details are updated at Customer SSN= "+Update_r, 0, Noc);
		}
	}
	/*
	 * Customer MailingAddressZipCode is mandatory and it is in specified format in 5 digits (for example 44444)
	 */
	@Test(priority=5)
	public void ArthaDB_MailingAddressZipCode_Check_TC005() throws SQLException
	{
		int Noc=0; 
		ResultSet rs=stmt.executeQuery("SELECT * FROM `arthadb`.`customers` where MailingAddressZipCode NOT REGEXP '[0-9]{5}'");
		List<String> Zip_rs=new ArrayList<String>();
		while(rs.next())
		{
			Noc=rs.getRow();
			Zip_rs.add(rs.getString("SSN"));
		}
		if(Noc==0)
		{
			Assert.assertEquals("Customer - MailingAddressZipCode check is Passed",0,Noc);
		}
		else
		{
			StringBuffer Zip_r=new StringBuffer();
			for(int k=0;k<Zip_rs.size();k++)
			{
				Zip_r.append(Zip_rs.get(k)+",");
			}
			Assert.assertEquals("Customer - MailingAddressZipCode check is Failed at SSN="+Zip_r, 0, Noc);
		}
	}
}
