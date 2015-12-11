package olx;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;

import autoitx4java.AutoItX;

import com.jacob.com.LibraryLoader;

class OLX {
  private static WebDriver PHdriver;
  static WebDriver driver;
  static String baseUrl;

  private static String PhotoPath;
  private static String PT;
  private static int photoCountValue;
  private static int lastRowNum;
  private static int startRowNum;
  private static int endRowNum;
  private static String excelPath="OLXSheet.xlsx";
  static  HashMap<String, String>  HM= new HashMap<String, String>();
 
  public static void main(String[] args) throws Exception {
	
      int lastRowNum=getLastRowNum(excelPath,"Sheet1");
      
     
      System.out.println(lastRowNum);
      for(int R=1;R<=lastRowNum;R++){
    	
    	  Thread.sleep(3000);
    	  getDataFromXL(excelPath, "Sheet1",R);     
//    		System.setProperty("webdriver.chrome.driver", "Browser\\chromedriver.exe");
    		driver = new FirefoxDriver();
    		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
    		driver.get("http://olx.in");
    		driver.findElement(By.xpath("//a[@id='topLoginLink']/span[2]/strong")).click();		    
    		String urlbeforeenterusername= driver.getCurrentUrl();
    		driver.findElement(By.id("userEmail")).clear();
    		driver.findElement(By.id("userEmail")).sendKeys(HM.get("Uname"));
    		driver.findElement(By.id("userPass")).clear();
    		driver.findElement(By.id("userPass")).sendKeys((String)HM.get("Pwd"));
    		driver.findElement(By.id("se_userLogin")).click();
    	    Thread.sleep(4000);
    		String urlafterentersubmit= driver.getCurrentUrl();
    		if(urlbeforeenterusername.matches(urlafterentersubmit)){
    		System.out.println("Either Uname or Password is wrong at Row No."+""+R);
    	     return;
    		}
    		
	        String startRow	=HM.get("RowStart");
	        startRowNum=Integer.parseInt(startRow);
	        String endRoW=HM.get("RowEnd");                    
    	    endRowNum=Integer.parseInt(endRoW);
    		if(startRowNum>endRowNum){
    		System.out.println("Pl check Value Enter in Sheet1 at Row,"+""+R+""+"RowStart Value is less than RowEnd Value"); 
    		 return;
    		}else if(startRowNum<=0){
    		 System.out.println("Pl check Value Enter in Sheet1 at Row,"+""+R+""+"RowStart Value cannot be"+""+0+"or less than Zero"); 
    		return; 
    		}
    			    	
	    	for(int Z =startRowNum;Z<=endRowNum;Z++){
	    		doPost(Z);
	    	}
	    	driver.quit();
	    	try{
	    		ChangeIP();	
	    	}catch(Exception e){
	    		
	    	}
	    	
       }	 
        
    	
    }
    	             //////////////                   
 

  public static void doPost(int Z) throws Exception{
	  getDataFromXL(excelPath, "Sheet2", Z);
	  String Condition=HM.get("Condition").trim().toUpperCase();
	  if(Condition.equalsIgnoreCase("Y")){
			  Thread.sleep(6000);
		       if(checkData(Z)==false){
		    	   return;
		       }
		      
	    	  driver.findElement(By.id("postNewAdLink")).click();
		      driver.findElement(By.id("add-title")).clear();
		      driver.findElement(By.id("add-title")).sendKeys((String)HM.get("Title"));
		      driver.findElement(By.xpath("//dl[@id='targetrenderSelect1-0']/dt/a")).click();
		      driver.findElement(By.xpath("//a[@id='cat-3']/span")).click();
		      String type =HM.get("Category");
		     if(type.equalsIgnoreCase("Apartments")){
		      driver.findElement(By.xpath("//div[@id='category-3']/div[2]/div[2]/div/ul/li[2]/a/span")).click();
		     }else if(type.equalsIgnoreCase("Houses")){
		  	   driver.findElement(By.xpath("//div[@id='category-3']/div[2]/div[2]/div/ul/li[1]/a/span")).click(); 
		     }
		  
		    String act= HM.get("Activity");
		    if(act.equalsIgnoreCase("HRent")){
		     driver.findElement(By.xpath("//div[@id='category-1309']//a/span[contains(text(),'Rent')]")).click();
		     }else if(act.equalsIgnoreCase("HSale")){
		        driver.findElement(By.xpath("//div[@id='category-1309']//a/span[contains(text(),'Sale')]")).click();
		     }
		    if(act.equalsIgnoreCase("ARent")){
		     driver.findElement(By.xpath("//div[@id='category-1307']//a/span[contains(text(),'Rent')]")).click();
		    }else if(act.equalsIgnoreCase("ASale")){   
		    driver.findElement(By.xpath("//div[@id='category-1307']//a/span[contains(text(),'Sale')]")).click();
		    }
		     driver.findElement(By.name("data[param_price][1]")).clear();
		     driver.findElement(By.name("data[param_price][1]")).sendKeys((String)HM.get("Price"));
		       String fur=HM.get("Furnished");
		      driver.findElement(By.xpath("//dl[@id='targetparam15']/dt/a")).click();
		      if(fur.equalsIgnoreCase("Yes")){
		      	driver.findElement(By.linkText("Yes")).click();
		      	
		      }else{
		      	driver.findElement(By.linkText("No")).click();
		     
		      }
		      
		      
		      String room=HM.get("Rooms");
		      driver.findElement(By.xpath("//dl[@id='targetparam17']/dt/a")).click();
		      if(room.equalsIgnoreCase("1BHK")){
		      	driver.findElement(By.linkText("1 room")).click();
		      	}
		      else if(room.equalsIgnoreCase("2BHK")){
		      	driver.findElement(By.linkText("2 rooms")).click();
		      
		      } else if(room.equalsIgnoreCase("3BHK")){
		      	driver.findElement(By.linkText("3 rooms")).click();
		      
		      }else{
		      	driver.findElement(By.linkText("4 and more")).click();
		      }
		      

		      driver.findElement(By.id("param325")).sendKeys((String)HM.get("SqrFeet"));
		      
		      driver.findElement(By.id("add-description")).clear();
		      driver.findElement(By.id("add-description")).sendKeys((String)HM.get("Description"));
		  
	

		           
		      String PhotoFolderPath ="OLXImagePost/RealEstate/"+HM.get("Category").toUpperCase()+"/"+Z;
		    //int photoCountValue=  (int) new File("").length();//????????????????????
		      File fileFolder=new File(PhotoFolderPath);
		      File[] arrFiles=fileFolder.listFiles();
		      
		    for(int p=0;p<=arrFiles.length-1;p++){
		              Thread.sleep(5000);	
		    
	    			     driver.findElement(By.xpath("//li[@id='add-file-"+(p+1)+"']")).click();;
	    			     
	    			      Thread.sleep(6000);
	    			            
	    			      File fileObj= arrFiles[p];
	    			      PhotoPath=fileObj.getAbsolutePath(); 
	    			      UploadFile(PhotoPath,"File Upload");
		       }

		    Thread.sleep(4000);
		     driver.findElement(By.id("add-person")).clear();
		     driver.findElement(By.id("add-person")).sendKeys((String)HM.get("AdPosterName"));
		     driver.findElement(By.id("add-phone")).clear();
		     driver.findElement(By.id("add-phone")).sendKeys((String)HM.get("Phone"));
		     driver.findElement(By.name("data[city]")).clear();
		     driver.findElement(By.name("data[city]")).sendKeys((String)HM.get("City"));
		     Thread.sleep(2000);
		     new Actions(driver).moveToElement(driver.findElement(By.xpath("//ul[@id='autosuggest-geo-ul']/li/a"))).build().perform();
		     driver.findElement(By.xpath("//ul[@id='autosuggest-geo-ul']/li/a")).click();
		     //driver.findElement(By.linkText((String)HM.get("City"))).click();  //Handle auto-suggestion of City element
		     driver.findElement(By.name("data[district]")).clear();
		     driver.findElement(By.name("data[district]")).sendKeys((String)HM.get("Locality"));
		     try{
		    	 driver.findElement(By.linkText((String)HM.get("Locality"))).click();
		     }catch(Exception e){
		    	 //e.printStackTrace();
		     }
		     
		   
		  
		    Thread.sleep(4000);
		    
		   driver.findElement(By.id("save")).click();
		  
		   Thread.sleep(4000);
		    driver.findElement(By.xpath("//a[@class='button br3 wide']")).click();
		    setStatusXL(excelPath, Z);    

	  }
  }
  static boolean checkData(int Z){
	  	if(HM.get("SqrFeet").isEmpty()==true){
	        System.out.println("Please Enter the Value in SqrFeet at Row Number"+""+Z);
	       return false;
		}
		if(HM.get("Furnished").isEmpty()==true){
		      System.out.println("Please Enter the Yes or No in Furnished at Row Number"+""+Z);
		   return false;
		}
		
		if(HM.get("Rooms").isEmpty()==true){
		       System.out.println("Please Enter the 1BHK or 2BHK or 3BHK or 4BHK  in Rooms at Row Number"+""+Z);
		       return false;
		}
		
		if(HM.get("Title").isEmpty()==true){
		   System.out.println("Please Enter the Title  at Row Number"+""+Z);
		   return false;
		}
		if(HM.get("Activity").isEmpty()==true){
		 System.out.println("Please Enter the Activity 'HRent' or 'HSale' or 'ARent' or 'ASale'  at Row Number"+""+Z);
		 return false;
		}
		
		if(HM.get("Category").isEmpty()==true){
			 System.out.println("Please Enter the Category 'Apartments' or 'Houses' at Row Number"+""+Z);
			 return false;
		}
		if(HM.get("City").isEmpty()==true){
			System.out.println("Please Enter the City  at Row Number"+""+Z);
			 return false;
		}
			
		if(HM.get("Price").isEmpty()==true){
			System.out.println("Please Enter the Price at Row Number"+""+Z);
			 return false;
		}
		if(HM.get("Phone").isEmpty()==true){
			System.out.println("Please Enter the Phone No. at Row Number"+""+Z);
			 return false;
		}
   return true;
  }
  
  
	  static void setStatusXL(String path, int rowNum) throws Exception {
	        	  FileInputStream fis= new FileInputStream(path);
		          Workbook wbook=WorkbookFactory.create(fis);
				  wbook.getSheet("Sheet2").getRow(rowNum).createCell(0).setCellValue("N");
				  FileOutputStream fos= new FileOutputStream(path);
				  wbook.write(fos);
				  fis.close();
				  fos.close();
	  }
   static void getDataFromXL(String Filepath,String SheetName,int rowNum) throws Exception {
	    FileInputStream fis=new FileInputStream(Filepath);
	    Workbook wBookObj=WorkbookFactory.create(fis);
	    Sheet sheetObj=wBookObj.getSheet(SheetName);
	    int lastRow=sheetObj.getLastRowNum();
	    lastRowNum= lastRow+1;
	    if(lastRowNum>rowNum){
		Row ValRow=sheetObj.getRow(rowNum);
		Row ValRow0=sheetObj.getRow(0);
		int lastCellnum=ValRow0.getLastCellNum();
		String KeyName="";
		String KeyVal="";
	
		for(int i=0;i<lastCellnum;i++){
			KeyName=ValRow0.getCell(i).getStringCellValue();
			KeyVal=ValRow.getCell(i).getStringCellValue();
			
			 HM.put(KeyName, KeyVal);
		}

	    	
	    }else{
	    	System.out.println("Data is not found or finished in "+SheetName+"at Row"+rowNum);
	    
	    }
		
    }
                    static int getLastRowNum(String Filepath,String SheetName) throws Exception{
	                         FileInputStream fis=new FileInputStream(Filepath);
	                           Workbook wBookObj=WorkbookFactory.create(fis);
	                               Sheet sheetObj=wBookObj.getSheet(SheetName);
	                                    int lastRow=sheetObj.getLastRowNum();
		                                             return lastRow;
	   }
                    
                    public static void ChangeIP() throws Exception {
                  	  setUp();
                  	  //checkIP();
                  	  DisConnect();
                  	  Thread.sleep(5000);
                  	  ConnectTo();
                  	  //checkIP();
                  	  PHdriver.quit();
                  		
                  		
                    }
                    
                    public static void setUp() throws Exception {
//                  	  System.setProperty("webdriver.chrome.driver", "Browser\\chromedriver.exe");
                  	  PHdriver = new FirefoxDriver();
                      	 
                  	  PHdriver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
                    }
                   
                    public static void ConnectTo() throws Exception {
                  	
                  	  PHdriver.get("http://tata-photon-max.wifi/html/home.html");
                      try{
                      	PHdriver.findElement(By.id("username")).clear();
                      	PHdriver.findElement(By.id("username")).sendKeys("admin");
                      	PHdriver.findElement(By.id("password")).clear();
                      	PHdriver.findElement(By.id("password")).sendKeys("admin");
                      	PHdriver.findElement(By.linkText("Log In")).click();
                  	    System.out.println("Login Done successfully");
                  	    
                        }catch(Exception e){
                  	
                        }
                    
                      try{
                      	PHdriver.findElement(By.id("connect_btn")).click();
                      	Thread.sleep(5000);
                      	PHdriver.findElement(By.id("disconnect_btn")).isDisplayed();
                  	   System.out.println("Internet connected successfully");
                      }catch(Exception e){
                      	System.out.println("Already connected");
                      	//driver.findElement(By.id("disconnect_btn")).click();	
                      }
                  	
                    

                    }
                    public static void DisConnect() throws Exception {
                  		
                  	  PHdriver.get("http://tata-photon-max.wifi/html/home.html");
                  	    try{
                  	    	PHdriver.findElement(By.id("username")).clear();
                  	    	PHdriver.findElement(By.id("username")).sendKeys("admin");
                  	    	PHdriver.findElement(By.id("password")).clear();
                  	    	PHdriver.findElement(By.id("password")).sendKeys("admin");
                  	    	PHdriver.findElement(By.linkText("Log In")).click();
                  	      }catch(Exception e){
                  		
                  	      }
                  	 	
                  	 	 try{
                  	 		 Thread.sleep(4000);
                  	 		PHdriver.findElement(By.id("disconnect_btn")).click();
                  	 		Thread.sleep(5000);
                  	 		PHdriver.findElement(By.id("connect_btn")).isDisplayed();
                  	 	   System.out.println("Internet Disconnected successfully");
                  	 	   
                  	 	    }catch(Exception e){
                  	 	    	System.out.println("Already Disconnected");
                  	 	    	
                  	 	    }

                  	  }
                  public static void checkIP() throws AWTException, InterruptedException{
                  	try{
                  		PHdriver.get("https://www.google.co.in");
                  		PHdriver.findElement(By.id("lst-ib")).clear();
                  		PHdriver.findElement(By.id("lst-ib")).sendKeys("My Ip");
                  	    Thread.sleep(3000);
                  	    new Robot().keyPress(KeyEvent.VK_ENTER);
                  	    String x = PHdriver.findElement(By.xpath("//div[@class='_h4c _rGd vk_h']")).getText();
                  	    System.out.println(x);
                  	}catch(Exception e){}
                  	
                     }
                  public static void UploadFile(String fileToUpload,String DialogTitle) throws InterruptedException{
              		String jacobDllVersionToUse;
              		String jvmBitVersion=System.getProperty("sun.arch.data.model");
              		if (jvmBitVersion.contains("32")){
              		jacobDllVersionToUse = "jacob-1.18-M2-x86.dll";
              		}
              		else {
              		jacobDllVersionToUse = "jacob-1.18-M2-x64.dll";
              		}
              		File file =new File("jar",jacobDllVersionToUse);
              		System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());
              		
              			fileToUpload=new File(fileToUpload).getAbsolutePath();
              			AutoItX x = new AutoItX();
              			x.winActivate(DialogTitle);
              			x.winWaitActive(DialogTitle);
              			x.ControlSetText(DialogTitle, "", "Edit1",fileToUpload) ;
              			Thread.sleep(1000);
              			x.controlClick(DialogTitle, "", "Button1") ;
              			Thread.sleep(1000);
              					
              		
              	}
              	
  
	  }
  

