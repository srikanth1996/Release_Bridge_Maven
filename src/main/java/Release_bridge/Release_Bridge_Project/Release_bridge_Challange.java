package Release_bridge.Release_Bridge_Project;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.URL;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import java.util.Date;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.remote.Command;
import org.openqa.selenium.remote.CommandExecutor;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.HttpCommandExecutor;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.remote.Response;
import org.openqa.selenium.remote.SessionId;
//import org.openqa.selenium.remote.http.W3CHttpCommandCodec;
//import org.openqa.selenium.remote.http.W3CHttpResponseCodec;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;



public class Release_bridge_Challange 
{
	static JavascriptExecutor js;
	/*public static RemoteWebDriver createDriverFromSession(final SessionId sessionId, URL command_executor)
	{
    CommandExecutor executor = new HttpCommandExecutor(command_executor) 
    {

    @Override
    public Response execute(Command command) throws IOException {
        Response response = null;
        if (command.getName() == "newSession") {
            response = new Response();
            response.setSessionId(sessionId.toString());
            response.setStatus(0);
            response.setValue(Collections.<String, String>emptyMap());

            try {
                Field commandCodec = null;
                commandCodec = this.getClass().getSuperclass().getDeclaredField("commandCodec");
                commandCodec.setAccessible(true);
                commandCodec.set(this, new W3CHttpCommandCodec());

                Field responseCodec = null;
                responseCodec = this.getClass().getSuperclass().getDeclaredField("responseCodec");
                responseCodec.setAccessible(true);
                responseCodec.set(this, new W3CHttpResponseCodec());
            } catch (NoSuchFieldException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }

        } else {
            response = super.execute(command);
        }
        return response;
    }
    };

    return new RemoteWebDriver(executor, new DesiredCapabilities());
}*/
	
	
	public static void main(String args[]) throws InterruptedException, IOException, ParseException
	{
		//System.setProperty("webdriver.chrome.driver","C:\\chromedriver.exe");
		//System.setProperty("webdriver.gecko.driver", "C:\\Program Files\\Mozilla Firefox\\firefox.exe");
		WebDriver driver = new ChromeDriver();
	   // HttpCommandExecutor executor = (HttpCommandExecutor) ((RemoteWebDriver) driver).getCommandExecutor();
	    //URL url = executor.getAddressOfRemoteServer();
	    //SessionId session_id = ((RemoteWebDriver) driver).getSessionId();
	    
	    
	    WebDriver driver2 = driver;//createDriverFromSession(session_id, url);
	    driver2.get("https://ibpm.cisco.com/ermo/rlc/login/UbQGjVTgzTnxEePt5N9tzw%5B%5B*/!STANDARD");
	   
	    driver2.manage().window().maximize();
	    
	    synchronized (driver2){
	    	driver2.wait(20000);
	    	};
	   // driver2.wait(20000);
	    driver2.findElement(By.xpath("//*[@id=\"userInput\"]")).click();
	    
	    System.out.println("succesfully cleared username");
	    
	    
	    FileReader reader=new FileReader("common.properties");  
	      
	    Properties p=new Properties();  
	    p.load(reader);  
	    
	    System.out.println(p.getProperty("Username"));  
	    System.out.println(p.getProperty("password"));  
	    
	    driver2.findElement(By.xpath("//*[@id=\"userInput\"]")).sendKeys(p.getProperty("Username").toString());
	    System.out.println("succesfully entered username");
	    
	    driver2.findElement(By.xpath("//*[@id=\"login-button\"]")).click();
	    
	    System.out.println("sucessfully clicked next button");
	    
	    driver2.manage().timeouts().implicitlyWait(100,TimeUnit.SECONDS) ;
	    WebElement element = driver2.findElement(By.xpath("//*[@id=\"passwordInput\"]"));
	    WebDriverWait wait = new WebDriverWait(driver, 20); //here, wait time is 20 seconds

	    wait.until(ExpectedConditions.visibilityOf(element));
	    element.click();
	    element.sendKeys(p.getProperty("password").toString());
	    System.out.println("successfully entered password");
	    
	    
	   
	    WebElement element1 = driver2.findElement(By.xpath("//*[@id=\"login-button\"]"));
	    element1.click();
	    
	   
	    synchronized (driver2){
	    	driver2.wait(50000);
	    	};
	   
	  // driver2.findElement(By.xpath("//a[contains(text(), 'What is this?')]")).click();
	  
	    //	 System.out.println("successfully opened two factor");
	    	// driver2.findElement(By.cssSelector("body")).click();

	    	 System.out.println("BEfore shifting to iframe");
	    	driver2.switchTo().frame(0);
	    	synchronized (driver2){
		    	driver2.wait(20000);
		    	};
	        driver2.findElement(By.xpath("//div[@id=\'auth_methods\']/fieldset/div[2]/button")).click();
	        System.out.println("successfully clicked call me button");
	   /* driver2.findElement(By.cssSelector("body")).sendKeys(Keys.TAB);
	    System.out.println("clicked TAB1");
	    driver2.findElement(By.cssSelector("body")).sendKeys(Keys.TAB);
	    driver2.findElement(By.cssSelector("body")).sendKeys(Keys.TAB);
	    driver2.findElement(By.cssSelector("body")).sendKeys(Keys.RETURN);*/
	    synchronized (driver2){
	    	driver2.wait(20000);
	    	};
	   System.out.println("Sussefully loggedIn");
	   
	 // String text =   driver2.findElement(By.xpath("/html/body/div/div/div[1]/div/form/div[1]/fieldset/h2")).getText();
	  //System.out.println(text); 
	  //System.out.println("successfully clicked TAB button");
	    /*
	    if(element2.size() > 0)
	    {
	        System.out.println("Found h2 header Downloads");
	        System.out.println("size of elemet is :"+element2.size());
	        element2.get(0).sendKeys(Keys.TAB);
	    }*/
	    /*
	    element1.sendKeys(Keys.TAB);
	    element1.sendKeys(Keys.TAB);
	    element1.sendKeys(Keys.ENTER);*/
	 /*   WebDriverWait wait1 = new WebDriverWait(driver, 20); //here, wait time is 20 seconds

	    wait1.until(ExpectedConditions.visibilityOf(element1));
	    element1.click();*/
	   
	   // driver2.findElement(By.id("passwordInput")).click();
	    synchronized (driver2){
	    	driver2.wait(20000);
	    	};
	    driver2.manage().window().setSize(new Dimension(1616, 916));
	    
	    {
  	      List<WebElement> elements = driver2.findElements(By.name("RBMenuCasePortal_pyDisplayHarness_24"));
  	      assert(elements.size() > 0);
  	  }
	    // 3 | click | name=RBMenuCasePortal_pyDisplayHarness_24 | 
	    driver2.findElement(By.name("RBMenuCasePortal_pyDisplayHarness_24")).click();
	    System.out.println("=====Successfully opened and clicked menu tab ");
	    
	    synchronized (driver2){
	    	driver2.wait(20000);
	    	};
	    	{
	    	      List<WebElement> elements = driver2.findElements(By.linkText("Reports"));
	    	      assert(elements.size() > 0);
	    	  }
	    // 4 | click | linkText=Reports | 
	    driver2.findElement(By.linkText("Reports")).click();
	    System.out.println("===== Successfully clicked Reports tab=====");
	    synchronized (driver2){
	    	driver2.wait(20000);
	    	};
	    	{
	    	      List<WebElement> elements = driver.findElements(By.linkText("RFC Report"));
	    	      assert(elements.size() > 0);
	    	 }
	    // 5 | click | linkText=RFC Report | 
	    driver2.findElement(By.linkText("RFC Report")).click();
	    System.out.println("===== Successfully clicked RFC Report link ");
	    synchronized (driver2){
	    	driver2.wait(20000);
	    	};
	    // 6 | click | css=#Tab4 > #TABANCHOR > #TABSPAN > #RULE_KEY td:nth-child(2) > span | 
	    	driver2.findElement(By.xpath("//span[@id=\'\']")).click();
	    System.out.println("====== Sucessfully clicked RFC tab ==");
	    synchronized (driver2){
	    	driver2.wait(20000);
	    	};
	    // 9 | selectFrame | index=0 | 
	    driver2.switchTo().frame(0);
	    
	    
	    driver2.findElement(By.id("pyLabel")).click();
	    WebElement dropdown3 = driver2.findElement(By.id("pyLabel"));
	    Select dropdown2 = new Select(dropdown3);
	      List<WebElement> allElements_before = dropdown2.getOptions();
	      System.out.println("size of drop down is :"+allElements_before.size());
	      ArrayList<String> DropdownString1 = new ArrayList();
		     for(WebElement ele1 :allElements_before)
		     {
		    	 //System.out.println(ele.getText());
		    	 DropdownString1.add(ele1.getText());
		     }
	     
		     XSSFWorkbook workbook = new XSSFWorkbook(); 
	         
		        //Create a blank sheet
		        XSSFSheet sheet = workbook.createSheet("All Data");
		       // FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
	  
		        int count=0;
		        
		        Map<String, Object[]> data = new TreeMap<String, Object[]>();
			       
	        	
	        	
	 for(int k=0;k<DropdownString1.size();k++)
	  	{
	  	
		 DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd");
			Date date = new Date();
			System.out.println("current Date is :"+dateFormat.format(date));
			String splitting[]=DropdownString1.get(k).toString().split(" ");
			String year = splitting[0].replace("FY", "20");
			String month = splitting[1].substring(0,3);
			String day = splitting[2];
			
			String month_new = null;
			if(month.equalsIgnoreCase("Jan")) {month_new="01";}
			if(month.equalsIgnoreCase("Feb")) {month_new="02";}
			if(month.equalsIgnoreCase("Mar")) {month_new="03";}
			if(month.equalsIgnoreCase("Apr")) {month_new="04";}
			if(month.equalsIgnoreCase("May")) {month_new="05";}
			if(month.equalsIgnoreCase("Jun")) {month_new="06";}
			if(month.equalsIgnoreCase("Jul")) {month_new="07";}
			if(month.equalsIgnoreCase("Aug")) {month_new="08";}
			if(month.equalsIgnoreCase("Sep")) {month_new="09";}
			if(month.equalsIgnoreCase("Oct")) {month_new="10";}
			if(month.equalsIgnoreCase("Nov")) {month_new="11";}
			if(month.equalsIgnoreCase("Dec")) {month_new="12";}
			
				
			String uidate = day+"/"+month_new+"/"+year;
			System.out.println("UI date is : "+uidate);
			Date date1=new SimpleDateFormat("dd/MM/yyyy").parse(uidate);  
			System.out.println("UI COnverted Date is :"+dateFormat.format(date1));
			
			
			 if(dateFormat.format(date).compareTo(dateFormat.format(date1)) < 0)
			 {
	    // 10 | click | id=pyLabel | d
	    driver2.findElement(By.id("pyLabel")).click();
	    System.out.println("======Sucessfully clicked dropdown label");
	    // 11 | select | id=pyLabel | label=FY20 March 08 Release Event
	    int index =0;
	    {
	      WebElement dropdown = driver2.findElement(By.id("pyLabel"));
	      Select dropdown1 = new Select(dropdown);
	      List<WebElement> allElements = dropdown1.getOptions();
	      System.out.println("Values present in Single Value Dropdown");
	    //  System.out.println("sixe of drop down :"+allElements.size()+" And element 0 is :"+allElements.get(0)+" And element 1 is :"+allElements.get(1));
	      
	      ArrayList<String> DropdownString = new ArrayList();
	     for(WebElement ele :allElements)
	     {
	    	 //System.out.println(ele.getText());
	    	 DropdownString.add(ele.getText());
	     }
	    
	     
	     
	     
	     	String option =DropdownString1.get(k).toString();
	     //	String xpath ="//option[. = 'FY20 March 08 Release Event']";
	     	String xpath1 = "//option[.= '"+option+"']";
	      dropdown.findElement(By.xpath(xpath1)).click();
	      for(int i=0;i<DropdownString.size();i++) 
	      {
	    	  if(DropdownString.get(i).toString().toLowerCase().equalsIgnoreCase(option))
	    	  {
	    		  index = i;
	    	  }
	    	  System.out.println("Array Lisy is :"+DropdownString.get(i));
	      }
	    }
	    
	    //Select dropdown1 = new Select(dropdown);
	    
	    int index1 = index+1;
	    // 12 | click | css=option:nth-child(48) | 
	 //   String optionselect1 = "option:nth-child(48)";
	    String optionselect  = "option:nth-child("+index1+")";
	    System.out.println("css selector path is :"+optionselect);
	    driver2.findElement(By.cssSelector(optionselect)).click();
	    System.out.println("=====Sucessfully selected : "+DropdownString1.get(k).toString());
	    js = (JavascriptExecutor) driver2;
	    // 13 | runScript | window.scrollTo(0,0) | 
	    js.executeScript("window.scrollTo(0,0)");
	    
	    synchronized (driver2){
	    	driver2.wait(30000);
	    	};
	    	
	    	
	    	//Appliyong PROF Filter
	    	 /*driver2.findElement(By.cssSelector(".cellCont:nth-child(6) span:nth-child(3) > #pui_filter")).click();
	    	 synchronized (driver2){
	 	    	driver2.wait(20000);
	 	    	};   
	    	 driver2.findElement(By.id("pySearchText")).click();
	    	 synchronized (driver2){
	 	    	driver2.wait(20000);
	 	    	};
	    	    driver2.findElement(By.id("pySearchText")).sendKeys("PROD");
	    	    synchronized (driver2){
	    	    	driver2.wait(20000);
	    	    	};
	    	    driver2.findElement(By.cssSelector("div > .pzhc:nth-child(1)")).click();
	    	    synchronized (driver2){
	    	    	driver2.wait(50000);
	    	    	};
	    	    	driver2.findElement(By.xpath("//table[@id=\'bodyTbl_right\']/tbody/tr/th/div/div/div")).click();
	    	    	synchronized (driver2){
		    	    	driver2.wait(50000);
		    	    	};*/
		    	    	
	    	List  rows = driver2.findElements(By.xpath(".//*[@id='gridBody_right']/table/tbody/tr/td[1]")); 
	        System.out.println("No of rows are : " + rows.size());
	        
	        
	        data = new TreeMap<String, Object[]>();
	        if(count==0)
	        {
	        	data.put("1", new Object[] {"Service Allignment", "ReleaseId","Release Name ","RFC Id"," Status "," RFC Title "," Submitted by "," Submitted Date "," RFC Start Date "," RFO End Date"});
	        }
	        for(int i=2;i<=rows.size()+1;i++)
	        {

	        	String getServiceAllignment = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[1]";
	        	String serviceAllignment = driver2.findElement(By.xpath(getServiceAllignment)).getText();
	        	System.out.println("Service Allignment :"+serviceAllignment);
	        	if(!serviceAllignment.equalsIgnoreCase("No items"))
	        	{
	        	System.out.println("===========================");
	        	String releaseId = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[2]";
	        	String getReleaseId = driver2.findElement(By.xpath(releaseId)).getText();
	        	System.out.println("Get Release Ids :"+getReleaseId);
	        	
	        	String releaseName = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[3]";
	        	String getreleaseName = driver2.findElement(By.xpath(releaseName)).getText();
	        	System.out.println("Get releaseName :"+getreleaseName);
	        	
	        	
	        	String rfcId = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[4]";
	        	String getrfcId = driver2.findElement(By.xpath(rfcId)).getText();
	        	System.out.println("Get rfcId Ids :"+getrfcId);
	        	
	        	
	        	String statusId = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[5]";
	        	String getstatusId = driver2.findElement(By.xpath(statusId)).getText();
	        	System.out.println("Get statusId Ids :"+getstatusId);
	        	
	        	String rfcTitle = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[6]";
	        	String getrfcTitle = driver2.findElement(By.xpath(rfcTitle)).getText();
	        	System.out.println("Get getrfcTitle Ids :"+getrfcTitle);
	        	
	        	String cecId = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[7]";
	        	String getcecId = driver2.findElement(By.xpath(cecId)).getText();
	        	System.out.println("Get getcecId Ids :"+getcecId);
	        	
	        	String submittedDate = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[8]";
	        	String getsubmittedDate = driver2.findElement(By.xpath(submittedDate)).getText();
	        	System.out.println("Get getsubmittedDate Ids :"+getsubmittedDate);
	        	
	        	String rfcstartDate = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[9]";
	        	String getrfcstartDate = driver2.findElement(By.xpath(rfcstartDate)).getText();
	        	System.out.println("Get getrfcstartDate Ids :"+getrfcstartDate);
	        	
	        	String rfcendDate = ".//*[@id='gridBody_right']/table/tbody/tr["+i+"]/td[10]";
	        	String getrfcendDate = driver2.findElement(By.xpath(rfcendDate)).getText();
	        	System.out.println("Get getrfcstartDate Ids :"+getrfcendDate);
	        	
	        	if(getrfcTitle.contains("PROD")||getrfcTitle.contains("PRD"))
	        	data.put(Integer.toString(i), new Object[] {serviceAllignment, getReleaseId,getreleaseName,getrfcId,getstatusId,getrfcTitle,getcecId,getsubmittedDate,getrfcstartDate,getrfcendDate});
	        	else
	        		continue;
	        }
	        	count++;
	        }
			 
	        Set<String> keyset = data.keySet();
	        int rowCount = sheet.getLastRowNum();
	        
	        System.out.println("Last rownum is :"+ rowCount);
	        //int rownum = 0;
	        for (String key : keyset)
	        {
	        	
	            Row row = sheet.createRow(++rowCount);
	        	
	            Object [] objArr = data.get(key);
	            int cellnum = 0;	
	            Cell cell = row.createCell(cellnum);
                cell.setCellValue(rowCount);
	            for (Object obj : objArr)
	            {
	                cell = row.createCell(++cellnum);
	               if(obj instanceof String)
	                    cell.setCellValue((String)obj);
	                else if(obj instanceof Integer)
	                    cell.setCellValue((Integer)obj);
	            }
	        }
	       // workbook.write(out);
	        
	  	}
	  	}
	        try
	        {
	            //Write the workbook in file system
	            FileOutputStream out = new FileOutputStream(new File("Release_Brdige_Prod_Release_items.xlsx"));
	            workbook.write(out);
	            out.close();
	            System.out.println("Release_Brdige_Prod_Release_items.xlsx written successfully on disk.");
	        } 
	        catch (Exception e) 
	        {
	            e.printStackTrace();
	        }
	        
	        
	        	//send mail to particular mails
	        send_emailWithAttachement();
	 
	    driver2.close();
	    

	}
	
	
	
	public static void send_emailWithAttachement() throws IOException
	   {
		FileReader reader1=new FileReader("common.properties");  
	      
	    Properties p1=new Properties();  
	    p1.load(reader1);  
		      // Recipient's email ID needs to be mentioned.
		      String to = p1.getProperty("mailTo").toString();
		      
		      // Sender's email ID needs to be mentioned
		      String from = p1.getProperty("mailFrom").toString();

		      final String username = p1.getProperty("Username").toString()+"@cisco.com";//change accordingly
		      final String password = p1.getProperty("password").toString();//change accordingly

		      // Assuming you are sending email through relay.jangosmtp.net
		      String host = "rcdn-mx-01.cisco.com";

		      Properties props = new Properties();
		      props.put("mail.smtp.auth", "true");
		      props.put("mail.smtp.starttls.enable", "true");
		      props.put("mail.smtp.host", host);
		      props.put("mail.smtp.port", "25");

		      // Get the Session object.
		      Session session = Session.getInstance(props,
		         new javax.mail.Authenticator() {
		            protected PasswordAuthentication getPasswordAuthentication() {
		               return new PasswordAuthentication(username, password);
		            }
		         });

		      try {
		         // Create a default MimeMessage object.
		         Message message = new MimeMessage(session);

		         // Set From: header field of the header.
		         message.setFrom(new InternetAddress(from));

		         // Set To: header field of the header.
		         message.setRecipients(Message.RecipientType.TO,
		            InternetAddress.parse(to));

		         // Set Subject: header field
		         message.setSubject("Release Bridge PROD/PRD Release Items");

		         // Create the message part
		         BodyPart messageBodyPart = new MimeBodyPart();

		         // Now set the actual message
		         messageBodyPart.setText("Hi , \n Please check for The following Excel for release PROD Data ");

		         // Create a multipar message
		         Multipart multipart = new MimeMultipart();

		         // Set text message part
		         multipart.addBodyPart(messageBodyPart);

		         // Part two is attachment
		         messageBodyPart = new MimeBodyPart();
		         String filename = "Release_Brdige_Prod_Release_items.xlsx";
		         DataSource source = new FileDataSource(filename);
		         messageBodyPart.setDataHandler(new DataHandler(source));
		         messageBodyPart.setFileName(filename);
		         multipart.addBodyPart(messageBodyPart);

		         // Send the complete message parts
		         message.setContent(multipart);

		         // Send message
		         Transport.send(message);

		         System.out.println("Sent message successfully....");
		  
		      } catch (MessagingException e) {
		         throw new RuntimeException(e);
		      }
		   
	   }
}
