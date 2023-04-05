package DataDrivenTest;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.edge.EdgeDriver;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openqa.selenium.support.ui.Select;


public class RegistrationTest {

	public static void main(String[] args) throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		System.setProperty("webdriver.edge.driver",
				"D:\\Programming Tools\\Eclipse IDE\\WebDriver\\Edge\\edgedriver_win64\\msedgedriver.exe"); // Set property for Edge Driver, change the path for your webdriver.exe
		WebDriver driver = new EdgeDriver();
		
		FileInputStream file = new FileInputStream("D:\\Projects\\Test Automation\\AutomationPractice\\AutomationPractice\\DataTest2.xlsx"); // Get the excel file
		XSSFWorkbook workBook = new XSSFWorkbook(file); // Get the workbook 
		XSSFSheet sheet = workBook.getSheet("Sheet1"); // Get the sheet inside of workbook

		int rowCount = sheet.getLastRowNum(); // 3 Rows
		
		Thread.sleep(5000);
		
		for (int i = 1; i <= rowCount; i++) { // Iterate and get the data for every cell
			XSSFRow currentRow = sheet.getRow(i); // Iterate the row

			String firstName = currentRow.getCell(0).getStringCellValue(); // Store the value of first name into firstName variable
			String lastName = currentRow.getCell(1).getStringCellValue(); // Store the value of last name into lastName variable
			String gender = currentRow.getCell(2).getStringCellValue(); // Store the value of gender into gender variable
			String hobbies = currentRow.getCell(3).getStringCellValue(); // Store the value of hobbies into hobbies variable
			String department = currentRow.getCell(4).getStringCellValue(); // Store the value of department into department variable
			String username = currentRow.getCell(5).getStringCellValue(); // Store the value of username into firstName variable
			String password = currentRow.getCell(6).getStringCellValue(); // Store the value of first name into username variable
			String confirmPassword = currentRow.getCell(7).getStringCellValue(); // Store the value of password into password variable
			String email = currentRow.getCell(8).getStringCellValue(); // Store the value of confirm password into confirmPassword variable
			String number = currentRow.getCell(9).getStringCellValue(); // Store the value of contact number into number variable
			String addInfo = currentRow.getCell(10).getStringCellValue(); // Store the value of addition information into addInfo variable

			driver.get("https://mytestingthoughts.com/Sample/home.html"); // Open the link of sample registration form
			driver.manage().window().maximize(); // Maximize the window

			byNameLocator(driver, firstName, "first_name", "First Name"); // Fill up the first name field
			byNameLocator(driver, lastName, "last_name", "Last Name"); // Fill up the last name field
			gender(driver, gender); // Select the gender
			hobbies(driver, hobbies); // Choos the hobbies
			department(driver, department); // Select the department
			byNameLocator(driver, username, "user_name", "Username"); // Fill up the username field
			byNameLocator(driver, password, "user_password", "Password"); // Fill up the password field
			byNameLocator(driver, confirmPassword, "confirm_password", "Confirm Password"); // Fill up the confirm password field
			byNameLocator(driver, email, "email", "Email"); // Fill up the email field
			byNameLocator(driver, number, "contact_no", "Contact Number"); // Fill up the number field
			byXpathLocator(driver, addInfo, "//textarea[@id='exampleFormControlTextarea1']", "Addition Info"); // Fill up the addition information field
			
			driver.findElement(By.xpath("//body/div[1]/form[1]/fieldset[1]/div[13]/div[1]/button[1]")).click(); // Submit the registration form
			System.out.println("Done ----- Submit");
			Thread.sleep(3000); // 3 secs delay for next data and registration 
			
			isSubmitted(driver); // Confirmation if the registration is successfully submitted
			System.out.println(); // New line before the next data

		}
		
		driver.quit(); // Close the window
		
	}

	private static void isSubmitted(WebDriver driver) { // Method for confirmation of registration
		String arMessage = driver.findElement(By.xpath("//div[@id='success_message']")).getText();
		String erMessage = "Success Success!.";
		if (arMessage.equals(erMessage)) {
			System.out.println("Successful !");
		} 
		else {
			System.out.println("Failed !");
		}
	}

	private static void department(WebDriver driver, String department) throws InterruptedException { // Method for selecting a department
		Select departmentDropdown = new Select(driver.findElement(By.name("department")));
		departmentDropdown.selectByVisibleText(department);
		System.out.println("Done ----- Department");
		Thread.sleep(1000);
	}

	private static void byXpathLocator(WebDriver driver, String value, String xpathLocator, String fieldName) throws InterruptedException { // Method for finding element using xpath locator
		driver.findElement(By.xpath(xpathLocator)).sendKeys(value);
		System.out.println("Done ----- " + fieldName);
		Thread.sleep(1000);
	}

	private static void byNameLocator(WebDriver driver, String value, String nameLocator, String fieldName) throws InterruptedException { // Method for finding element using name locator
		driver.findElement(By.name(nameLocator)).sendKeys(value);
		System.out.println("Done ----- " + fieldName);
		Thread.sleep(1000);
	}

	private static void hobbies(WebDriver driver, String hobbies) throws InterruptedException { // Method for choosing hobbies
		if (hobbies.equals("Reading")) {
			driver.findElement(By.xpath("/html[1]/body[1]/div[1]/form[1]/fieldset[1]/div[4]/div[1]/select[1]/option[1]")).click();
		} 
		else if (hobbies.equals("Singing")) {
			driver.findElement(By.xpath("/html[1]/body[1]/div[1]/form[1]/fieldset[1]/div[4]/div[1]/select[1]/option[2]")).click();
		} 
		else if (hobbies.equals("Swimming")){
			driver.findElement(By.xpath("/html[1]/body[1]/div[1]/form[1]/fieldset[1]/div[4]/div[1]/select[1]/option[3]")).click();
		} 
		else if (hobbies.equals("Running")){
			driver.findElement(By.xpath("/html[1]/body[1]/div[1]/form[1]/fieldset[1]/div[4]/div[1]/select[1]/option[4]")).click();
		}
		else if (hobbies.equals("Jogging")){
			driver.findElement(By.xpath("/html[1]/body[1]/div[1]/form[1]/fieldset[1]/div[4]/div[1]/select[1]/option[5]")).click();
		}
		else {
			System.out.println("Failed ----- Hobbies");
		}
		System.out.println("Done ----- Hobbies");
		Thread.sleep(1000);
	}

	private static void gender(WebDriver driver, String gender) throws InterruptedException { // Method for selecting gender
		String male = "Male";
		String female = "Female";
		if (gender.equals(male)) {
			driver.findElement(By.xpath("//input[@id='inlineRadioMale']")).click();
			System.out.println("Done ----- Gender");
		} 
		else if (gender.equals(female)) {
			driver.findElement(By.xpath("//input[@id='inlineRadioFemale']")).click();
			System.out.println("Done ----- Gender");
		}
		else {
			System.out.println("Failed ----- Gender");
		}
		Thread.sleep(1000);
	}
}
