package cn.tjuscs.cn;


import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.*;
import static org.junit.Assert.*;
import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;


public class TestAddress {
	private static String githubAddress;
	WebDriver driver;
	String baseUrl;
	@Before
	public void setUp() throws Exception{
		String driverPath = "D:\\java_project\\lab2\\geckodriver.exe";
		System.setProperty("webdriver.gecko.driver", driverPath);
		driver = new FirefoxDriver();
		baseUrl = "http://121.193.130.195:8800";
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	}
	
	@Test
	public void test() throws IOException {
		FileInputStream excel = new FileInputStream("D:\\java_project\\lab2\\软件测试名单.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(excel);
		excel.close();
		XSSFSheet xssfSheet = workbook.getSheetAt(0);
		driver.get(baseUrl + "/");
		for (int rowNum = 2; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {					
			XSSFRow xssfRow = xssfSheet.getRow(rowNum);	
			if (xssfRow != null) {						
				// 读取学号数据															
				double cellValue = xssfRow.getCell(1).getNumericCellValue(); 
				String num = new DecimalFormat("#").format(cellValue); 	
				
				// 读取github数据	
				XSSFCell githubAdd = xssfRow.getCell(3);	
				//address：excel中的地址
				String address = String.valueOf(githubAdd.getStringCellValue());
				
				//定义用于登陆账号id和密码password
				String id = num;
				String password = num.substring(num.length()-6);
				WebElement we = driver.findElement(By.name("id"));
				we.click();
				//找到对应组件，输入账号密码
				driver.findElement(By.name("id")).sendKeys(id);
				driver.findElement(By.name("password")).sendKeys(password);
				driver.findElement(By.id("btn_login")).click();
				//获取网页中的GitHub地址
				githubAddress = driver.findElement(By.id("student-git")).getText();
				//退回到初始的登陆界面，便于寻找下一个地址
				driver.findElement(By.id("btn_logout")).click();
				driver.findElement(By.id("btn_return")).click();
				
				assertEquals(githubAddress, address);
				
			}	
		}
	}
}

