package NetShipPackage;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Random;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class NetShipAll extends BaseInit {

	public static String ScrSht;
	public static String SheetMessage, SheetMessage1;
	// Staging
	public static String CustomerNameNSPL = "TEST CUSTOMER 950024 - #950024",
			CustomerNameSPL = "TEST CUSTOMER 950025 - #950025";
	public static String PartF1 = "PART 95002501", FSLName1 = "PJD - TEST FSL FOR 34691 F5020 (F5020)";

	@Test
	public void ForgetPassword() throws Exception {
		WebDriverWait wait = new WebDriverWait(Driver, 50);
		Driver.findElement(By.id("divUsername")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("hrefLogout")));
		Driver.findElement(By.id("hrefLogout")).click();
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("hrefForgotPW")));
		Driver.findElement(By.id("hrefForgotPW")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("forgotpwForm")));
		Driver.findElement(By.id("btnBckLogin")).click();
		wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.name("loginForm")));
		Driver.findElement(By.id("hrefForgotPW")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("forgotpwForm")));
		Driver.findElement(By.id("inputUsername")).clear();
		Driver.findElement(By.id("inputEmilId")).clear();
		Driver.findElement(By.id("btnFRPw")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("lblValidate")));

		String Message1 = Driver.findElement(By.id("lblValidate")).getText();

		if (Message1.contains(
				"Please refer to the following error(s) : - User Name cannot be blank. - Customer Code cannot be blank.")) {
			System.out.println("*******Validation is display Proper.*******");
			System.out.println("*******" + Message1 + "*******");
		} else {
			System.out.println("*******Validation is not display Proper.*******");
			System.out.println("*******" + Message1 + "*******");
		}

		Driver.findElement(By.id("inputUsername")).clear();
		Driver.findElement(By.id("inputUsername")).sendKeys("automation");
		Driver.findElement(By.id("inputEmilId")).clear();
		Driver.findElement(By.id("btnFRPw")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("lblValidate")));

		String Message2 = Driver.findElement(By.id("lblValidate")).getText();

		if (Message2.contains("Please refer to the following error(s) : - Customer Code cannot be blank.")) {
			System.out.println("*******Validation is display Proper.*******");
			System.out.println("*******" + Message2 + "*******");
		} else {
			System.out.println("*******Validation is not display Proper.*******");
			System.out.println("*******" + Message2 + "*******");
		}

		Driver.findElement(By.id("inputUsername")).clear();
		Driver.findElement(By.id("inputEmilId")).clear();
		Driver.findElement(By.id("inputEmilId")).sendKeys("Auto@123");
		Driver.findElement(By.id("btnFRPw")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("lblValidate")));

		String Message3 = Driver.findElement(By.id("lblValidate")).getText();

		if (Message3.contains("Please refer to the following error(s) : - User Name cannot be blank.")) {
			System.out.println("*******Validation is display Proper.*******");
			System.out.println("*******" + Message3 + "*******");
		} else {
			System.out.println("*******Validation is not display Proper.*******");
			System.out.println("*******" + Message3 + "*******");
		}

		Driver.findElement(By.id("inputUsername")).clear();
		Driver.findElement(By.id("inputUsername")).sendKeys("Autotest");
		Driver.findElement(By.id("inputEmilId")).clear();
		Driver.findElement(By.id("inputEmilId")).sendKeys("ravina.prajapati@samyak.com");
		Driver.findElement(By.id("btnFRPw")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("lblValidate")));

		String Message4 = Driver.findElement(By.id("lblValidate")).getText();

		if (Message4.contains(
				"A temporary password has been issued to the email address on file. Please log in and reset your password.")) {
			System.out.println("*******Validation is display Proper.*******");
			System.out.println("*******" + Message4 + "*******");
		} else {
			System.out.println("*******Validation is not display Proper.*******");
			System.out.println("*******" + Message4 + "*******");
		}
		Login();
	}

	@Test
	public void ActiveOrder() throws Exception {
		WebDriverWait wait = new WebDriverWait(Driver, 50);
		System.out.println("**********Active Order**********");
		Robot robot = new Robot();
		// Read data from Excel
		// DEV
		// File src0 = new File("./DataFiles/NetShipActiveOrderDEV.xlsx");
		// Staging
		File src0 = new File("./DataFiles/NetShipActiveOrderSTG.xlsx");
		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("ActiveOrder");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();

		// --Group By dropdown
		Select SelectSort2 = new Select(Driver.findElement(By.id("drpGrouping")));
		SelectSort2.selectByIndex(1);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		String SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		String value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(2);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(3);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(4);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(5);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(6);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(7);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(8);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(9);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));
		SelectedOp = SelectSort2.getFirstSelectedOption().getText();
		value = Driver.findElement(By.xpath("//h4[@class=\"ng-binding\"]")).getText();
		System.out.println("selected option is==" + SelectedOp);
		System.out.println("Value of " + SelectedOp + " option is==" + value);

		SelectSort2.selectByIndex(0);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//h4[@class=\"ng-binding\"]")));

		// --Sort By dropdown
		Select SelectSort3 = new Select(Driver.findElement(By.id("drpSorting")));
		SelectSort3.selectByIndex(1);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort3.selectByIndex(2);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort3.selectByIndex(0);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));
		Thread.sleep(2000);
		// --Row Count of Excel sheet
		System.out.println("Row Count ====> " + rcount);

		// --Checking Memo, Print and job for All rows of excel
		for (int i = 1; i <= rcount; i++) {
			System.out.println("\n********************************************************************************");
			System.out.println("\nJob ID ==> " + formatter.formatCellValue(sh0.getRow(i).getCell(0)));

			String MeJob = "idmemo_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));
			String PrJob = "idprint_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));
			String SeJob = "PickupId_N" + formatter.formatCellValue(sh0.getRow(i).getCell(1));

			// --Click on Memo
			// try {
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(MeJob)));
			Driver.findElement(By.id(MeJob)).click();
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

			// Driver.findElement(By.id("hlkBackToScreen")).click();
			Driver.navigate().back();
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ActiveOrderGrd")));
			/*
			 * } catch (Exception e) { System.out.
			 * println("There is no Memo Added in Job, Please add Memo first in this Job : "
			 * + formatter.formatCellValue(sh0.getRow(i).getCell(0))); }
			 */

			// --Click on Print button
			try {
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(PrJob)));
				Driver.findElement(By.id(PrJob)).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

				// --Transfer to new window
				String winHandleBefore = Driver.getWindowHandle();
				for (String winHandle : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle);
					Thread.sleep(2000);
				}

				Driver.close();

				// Switch back to original browser (first window)
				Driver.switchTo().window(winHandleBefore);
				Thread.sleep(2000);

			} catch (Exception e) {
				System.out.println("Print Label is not able to work on Click, Please check Manualy with this Job : "
						+ formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			}

			// -Search with PickUpID
			System.out.println(SeJob);
			List<WebElement> dynamicElement = Driver.findElements(By.id(SeJob));
			if (dynamicElement.size() != 0) {
				Driver.findElement(By.id(SeJob)).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

				Driver.navigate().back();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ActiveOrderGrd")));
			} else {
				System.out.println("*********YOUR JOB IS NOT DISPLAY IN LIST.*********");
			}

			System.out.println("\n********************************************************************************");
			System.out.println("\n********************************************************************************");

			Driver.findElement(By.id("txtGlobalSearch")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));
			Driver.findElement(By.id("txtGlobalSearch")).sendKeys(Keys.ENTER);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

			try {
				// --Click on Email
				Driver.findElement(By.id("hlkEmail")).click();
				wait.until(ExpectedConditions
						.visibilityOfAllElementsLocatedBy(By.xpath("//*[@class=\"ngdialog-content\"]")));

				Driver.findElement(By.id("txtEmail")).clear();
				Driver.findElement(By.id("btnSendEmail")).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("hrefErr")));

				String Message1 = Driver.findElement(By.id("hrefErr")).getText();

				if (Message1.contains("Please Enter Email.")) {
					System.out.println("\n*******Validation are display Proper.*******");
					System.out.println("*******" + Message1 + "*******");
				} else {
					System.out.println("*******Validation are not display Proper.*******");
					System.out.println("*******" + Message1 + "*******");
				}

				// --Send email
				Driver.findElement(By.id("txtEmail")).sendKeys("Ravina.prajapati@samyak.com");
				Driver.findElement(By.id("btnSendEmail")).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

				String Message2 = Driver.findElement(By.xpath("//*[@ng-bind=\"emailMessage\"]")).getText();

				if (Message2.equals("Email successfully sent!")) {
					System.out.println("*******Message is display Proper.*******");
					System.out.println("*******" + Message2 + "*******");
				} else {
					System.out.println("*******Validation is not display Proper.*******");
					System.out.println("*******" + Message2 + "*******");
				}
				// --Close Email popup
				Driver.findElement(By.id("btnclose")).click();
				Thread.sleep(2000);

				// --Print button
				Driver.findElement(By.id("idpdfprint")).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

				String winHandleBefore = Driver.getWindowHandle();
				for (String winHandle : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle);
					Thread.sleep(2000);
				}

				Driver.close();

				// Switch back to original browser (first window)
				Driver.switchTo().window(winHandleBefore);
				Thread.sleep(2000);

				Driver.findElement(By.id("btshipdetail")).click();
				Thread.sleep(15000);

				System.out.println("Bill of Lading is working proper.");
			} catch (Exception e) {
				System.out.println("Bill of Lading is not enable.");
				System.out.println("Please Process CSR Acknowledge OR TC Acknowledge to enable Bill of Lading.");
			}

			// Print BOL on Shipment Detail Screen=Not available
			/*
			 * try { Driver.findElement(By.id("idprintbol")).click(); Thread.sleep(5000);
			 * 
			 * String winHandleBefore1 = Driver.getWindowHandle(); for (String winHandle1 :
			 * Driver.getWindowHandles()) { Driver.switchTo().window(winHandle1); }
			 * Thread.sleep(15000);
			 * 
			 * Driver.close(); Thread.sleep(15000);
			 * Driver.switchTo().window(winHandleBefore1);
			 * 
			 * System.out.println("Print BOL is working proper."); } catch (Exception e) {
			 * System.out.println("Print BOL is not enable."); System.out.
			 * println("Please Process CSR Acknowledge OR TC Acknowledge to enable Print BOL."
			 * ); Thread.sleep(5000); }
			 * 
			 * // FAX BOL on Shipment Detail Screen=Not available try {
			 * Driver.findElement(By.id("hlkFax")).click(); Thread.sleep(5000);
			 * Driver.findElement(By.id("btnSendFAXBOL")).click(); Thread.sleep(15000);
			 * 
			 * String Message4 = Driver.findElement(By.id("errorid")).getText();
			 * Thread.sleep(5000);
			 * 
			 * if (Message4.equals("Please select atleast one Fax #s.")) {
			 * System.out.println("*******Validation Message is display Proper.*******");
			 * System.out.println("*******" + Message4 + "*******"); } else {
			 * System.out.println("*******Validation Message is not display Proper.*******"
			 * ); System.out.println("*******" + Message4 + "*******"); }
			 * Thread.sleep(5000);
			 * 
			 * robot.keyPress(KeyEvent.VK_ESCAPE); Thread.sleep(5000);
			 * 
			 * System.out.println("FAX BOL is working proper."); } catch (Exception e) {
			 * System.out.println("FAX BOL is not enable."); System.out.
			 * println("Please Process CSR Acknowledge OR TC Acknowledge to enable FAX BOL."
			 * ); Thread.sleep(5000); }
			 */

			// --Memo
			try {
				Driver.findElement(By.id("idmemogreen")).click();
				wait.until(ExpectedConditions
						.visibilityOfAllElementsLocatedBy(By.xpath("//*[@class=\"ngdialog-content\"]")));

				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(1000);

				System.out.println("Memo is working proper.");
			} catch (Exception e) {
				System.out.println("Memo is not enable.");
				System.out.println("Please ADD Memo from Connect to enable Memo.");
			}

			// --Charges
			try {
				Driver.findElement(By.id("idcharges")).click();
				wait.until(ExpectedConditions
						.visibilityOfAllElementsLocatedBy(By.xpath("//*[@class=\"ngdialog-content\"]")));
				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(1000);

				System.out.println("View Charges is working proper.");
			} catch (Exception e) {
				System.out.println("Please Check Check Box from Connect to view Charges in Net Ship.");
			}

			try {
				Driver.findElement(By.id("idupload")).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));
				wait.until(ExpectedConditions
						.visibilityOfAllElementsLocatedBy(By.xpath("//*[@ng-form=\"DocDetailsForm\"]")));
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("txtDocName")));
				Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
				// DEV
				// Driver.findElement(By.id("file")).sendKeys("./DataFiles/Job Upload Doc
				// DEV.xls");
				// Staging
				Driver.findElement(By.id("file")).sendKeys("./DataFiles/Job Upload Doc STG.xls");
				Thread.sleep(2000);
				Driver.findElement(By.id("btnUpload")).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("successid")));

				String Message6 = Driver.findElement(By.id("successid")).getText();

				if (Message6.equals("Upload/Import Process Completed !")) {
					System.out.println(Message6);
				} else {
					Message6 = Driver.findElement(By.id("errorid")).getText();
					SheetMessage = "*****Import Process is not Completed !*****";
					System.out.println(Message6);
				}

				Driver.findElement(By.id("btnSave")).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

				System.out.println("Upload link is working proper.");

				if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
					Driver.findElement(By.id("btnCancel")).click();
					Thread.sleep(2000);
				} else {
					Driver.findElement(By.id("btnSave")).click();
					wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

					Driver.findElement(By.id("btnsaveandclose")).click();
					Thread.sleep(2000);
				}
			} catch (Exception e) {
				try {
					Driver.findElement(By.id("iduploadgreen")).click();
					Thread.sleep(20000);

					Driver.findElement(By.id("hlkaddUpload")).click();
					Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
					// DEV
					// Driver.findElement(By.id("file")).sendKeys("./DataFiles/Job Upload Doc
					// DEV.xls");
					// Staging
					Driver.findElement(By.id("file")).sendKeys("./DataFiles/Job Upload Doc STG.xls");
					Thread.sleep(15000);
					Driver.findElement(By.id("btnUpload")).click();
					Thread.sleep(30000);

					String Message6 = Driver.findElement(By.id("successid")).getText();

					if (Message6.equals("Import Process Completed !")) {
						Message6 = "*****Import Process is Completed !*****";
						System.out.println(Message6);
						Thread.sleep(5000);
					} else {
						Message6 = Driver.findElement(By.id("errorid")).getText();
						SheetMessage = "*****Import Process is not Completed !*****";
						System.out.println(Message6);
						Thread.sleep(5000);
					}
					Thread.sleep(5000);

					Driver.findElement(By.id("btnSave")).click();
					Thread.sleep(5000);

					System.out.println("Upload link is working proper.");

					if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
						Driver.findElement(By.id("btnCancel")).click();
						Thread.sleep(15000);
					} else {
						Driver.findElement(By.id("btnSave")).click();
						Thread.sleep(15000);

						Driver.findElement(By.id("btnsaveandclose")).click();
						Thread.sleep(15000);
					}
				} catch (Exception f) {
					try {
						Driver.findElement(By.id("hlkUploadDocument")).click();
						Thread.sleep(20000);

						Driver.findElement(By.id("hlkaddUpload")).click();
						Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
						// DEV
						// Driver.findElement(By.id("file")).sendKeys("./DataFiles/Job Upload Doc
						// DEV.xls");
						// Staging
						Driver.findElement(By.id("file")).sendKeys("./DataFiles/Job Upload Doc STG.xls");
						Thread.sleep(15000);
						Driver.findElement(By.id("btnUpload")).click();
						Thread.sleep(30000);

						String Message6 = Driver.findElement(By.id("successid")).getText();

						if (Message6.equals("Import Process Completed !")) {
							Message6 = "*****Import Process is Completed !*****";
							System.out.println(Message6);
							Thread.sleep(5000);
						} else {
							Message6 = Driver.findElement(By.id("errorid")).getText();
							SheetMessage = "*****Import Process is not Completed !*****";
							System.out.println(Message6);
							Thread.sleep(5000);
						}
						Thread.sleep(5000);

						Driver.findElement(By.id("btnSave")).click();
						Thread.sleep(5000);

						System.out.println("Upload link is working proper.");

						if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
							Driver.findElement(By.id("btnCancel")).click();
							Thread.sleep(15000);
						} else {
							Driver.findElement(By.id("btnSave")).click();
							Thread.sleep(15000);

							Driver.findElement(By.id("btnsaveandclose")).click();
							Thread.sleep(15000);
						}
					} catch (Exception g) {
						System.out.println("There is no Upload Image display Please check Job manualy.");
					}
				}
			}

			// Click on Watch list
			try {
				Driver.findElement(By.id("watchListBlack")).click();
				Thread.sleep(15000);
				System.out.println("Watch list is working proper.");
			} catch (Exception e) {
				try {
					Driver.findElement(By.id("watchListGreen")).click();
					Thread.sleep(15000);
					System.out.println("Watch list is working proper.");
				} catch (Exception f) {
					System.out.println("There is no Watch List Image display Please check Job manualy.");
				}
			}

			try {
				Driver.findElement(By.id("hlkRefresh")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				System.out.println("There is no job on Active Order from Test Data Excel.");
				Thread.sleep(15000);
			}

			try {
				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				System.out.println("There is no job on Active Order from Test Data Excel.");
				Thread.sleep(15000);
			}
		}
	}

	@Test
	public void RecentDelivery() throws Exception {
		System.out.println("**********Recent Deliveries**********");
		Robot robot = new Robot();
		WebDriverWait wait = new WebDriverWait(Driver, 50);

		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipRecentDeliveredDEV.xlsx");
		// Staging
		File src0 = new File(".\\DataFiles\\NetShipRecentDeliveredSTG.xlsx");

		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("RecentDelivered");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();

		System.out.println("Row Count ====> " + rcount);

		Driver.findElement(By.id("RecentDeliveredTab")).click();
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		Select SelectSort0 = new Select(Driver.findElement(By.id("drpcustomer")));
		SelectSort0.selectByIndex(1);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

//		SelectSort0.selectByIndex(2);
//		Thread.sleep(15000);
//		
//		SelectSort0.selectByIndex(3);
//		Thread.sleep(15000);
//		
//		SelectSort0.selectByIndex(4);
//		Thread.sleep(15000);

		SelectSort0.selectByIndex(0);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		Select SelectSort1 = new Select(Driver.findElement(By.id("drpHours")));
		SelectSort1.selectByIndex(1);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort1.selectByIndex(2);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort1.selectByIndex(3);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort1.selectByIndex(0);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort1.selectByIndex(3);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		Select SelectSort2 = new Select(Driver.findElement(By.id("drpGrouping")));
		SelectSort2.selectByIndex(1);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(2);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(3);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(4);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(5);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(6);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(7);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(8);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(9);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort2.selectByIndex(0);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		Select SelectSort3 = new Select(Driver.findElement(By.id("drpSorting")));
		SelectSort3.selectByIndex(1);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort3.selectByIndex(2);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		SelectSort3.selectByIndex(0);
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@class=\"ajax-loadernew\"]")));

		for (int i = 1; i <= rcount; i++) {
			System.out.println("\n********************************************************************************");
			System.out.println("\nJob ID ==> " + formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			Thread.sleep(5000);

			String MeJob = "idmemo_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));
			String PrJob = "idprint_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));
			String SeJob = "JobId_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));

			try {
				Driver.findElement(By.id(MeJob)).click();
				Thread.sleep(20000);

				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("There is no Memo Added in Job, Please add Memo first in this Job : "
						+ formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			}

			try {
				Driver.findElement(By.id(PrJob)).click();
				Thread.sleep(5000);

				String winHandleBefore2 = Driver.getWindowHandle();
				for (String winHandle2 : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle2);
				}
				Thread.sleep(15000);

				Driver.close();
				Thread.sleep(15000);
				Driver.switchTo().window(winHandleBefore2);
			} catch (Exception e) {
				System.out.println("Print Label is not able to work on Click, Please check Manualy with this Job : "
						+ formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			}

			List<WebElement> dynamicElement = Driver.findElements(By.id(SeJob));
			if (dynamicElement.size() != 0) {
				Driver.findElement(By.id(SeJob)).click();
				Thread.sleep(20000);

				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(5000);
			} else {
				System.out.println("*********YOUR JOB IS NOT DISPLAY IN LIST.*********");
			}

			System.out.println("\n********************************************************************************");
			System.out.println("\n********************************************************************************");

			Driver.findElement(By.id("txtGlobalSearch")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));
			Driver.findElement(By.id("txtGlobalSearch")).sendKeys(Keys.ENTER);
			Thread.sleep(20000);

			try {
				Driver.findElement(By.id("hlkBillofLading")).click();
				Thread.sleep(5000);

				Driver.findElement(By.id("btemailbol")).click();
				Thread.sleep(5000);

				Driver.findElement(By.id("txtfrom")).clear();
				Driver.findElement(By.id("btnSendBOL")).click();
				Thread.sleep(5000);

				String Message1 = Driver.findElement(By.id("idValidation")).getText();

				if (Message1.contains("Address is required.")) {
					System.out.println("\n*******Validation are display Proper.*******");
					System.out.println("*******" + Message1 + "*******");
					Thread.sleep(5000);
				} else {
					System.out.println("*******Validation are not display Proper.*******");
					System.out.println("*******" + Message1 + "*******");
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("txtfrom")).sendKeys("pdoshi@samyak.com");
				Driver.findElement(By.id("txtto")).sendKeys("puagent@samyakinfo.com,dlagent@samyakinfo.com");
				Driver.findElement(By.id("btnSendBOL")).click();
				Thread.sleep(5000);
				Driver.findElement(By.xpath("/html/body/div[6]/div/div/div[3]/button[1]")).click();
				Thread.sleep(15000);

				String Message2 = Driver.findElement(By.id("saveid")).getText();

				if (Message2.equals("B.O.L # successfully send via Email")) {
					System.out.println("*******Message is display Proper.*******");
					System.out.println("*******" + Message2 + "*******");
				} else {
					System.out.println("*******Validation is not display Proper.*******");
					System.out.println("*******" + Message2 + "*******");
				}
				Thread.sleep(5000);

				Driver.findElement(By.id("btfaxbol")).click();
				Thread.sleep(5000);
				Driver.findElement(By.id("btnSendFAX")).click();
				Thread.sleep(5000);

				String Message3 = Driver.findElement(By.id("errorid")).getText();

				if (Message3.equals("Please select atleast one Fax #s.")) {
					System.out.println("*******Message is display Proper.*******");
					System.out.println("*******" + Message3 + "*******");
				} else {
					System.out.println("*******Validation is not display Proper.*******");
					System.out.println("*******" + Message3 + "*******");
				}
				Thread.sleep(5000);

				// FAX Setup Link
				// Driver.findElement(By.id("hrefFax")).click();
				// Thread.sleep(15000);

				Driver.findElement(By.id("btprint")).click();
				Thread.sleep(5000);

				String winHandleBefore = Driver.getWindowHandle();
				for (String winHandle : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle);
				}
				Thread.sleep(15000);

				Driver.close();
				Thread.sleep(15000);
				Driver.switchTo().window(winHandleBefore);

				Driver.findElement(By.id("btshipdetail")).click();
				Thread.sleep(15000);

				System.out.println("Bill of Lading is working proper.");
			} catch (Exception e) {
				System.out.println("Bill of Lading is not enable.");
				System.out.println("Please Process CSR Acknowledge OR TC Acknowledge to enable Bill of Lading.");
			}

			// Print BOL on Shipment Detail Screen
			try {
				Driver.findElement(By.id("idprintbol")).click();
				Thread.sleep(5000);

				String winHandleBefore1 = Driver.getWindowHandle();
				for (String winHandle1 : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle1);
				}
				Thread.sleep(15000);

				Driver.close();
				Thread.sleep(15000);
				Driver.switchTo().window(winHandleBefore1);

				System.out.println("Print BOL is working proper.");
			} catch (Exception e) {
				System.out.println("Print BOL is not enable.");
				System.out.println("Please Process CSR Acknowledge OR TC Acknowledge to enable Print BOL.");
				Thread.sleep(5000);
			}

			// FAX BOL on Shipment Detail Screen
			try {
				Driver.findElement(By.id("hlkFax")).click();
				Thread.sleep(5000);
				Driver.findElement(By.id("btnSendFAXBOL")).click();
				Thread.sleep(15000);

				String Message4 = Driver.findElement(By.id("errorid")).getText();
				Thread.sleep(5000);

				if (Message4.equals("Please select atleast one Fax #s.")) {
					System.out.println("*******Validation Message is display Proper.*******");
					System.out.println("*******" + Message4 + "*******");
				} else {
					System.out.println("*******Validation Message is not display Proper.*******");
					System.out.println("*******" + Message4 + "*******");
				}
				Thread.sleep(5000);

				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(5000);

				System.out.println("FAX BOL is working proper.");
			} catch (Exception e) {
				System.out.println("FAX BOL is not enable.");
				System.out.println("Please Process CSR Acknowledge OR TC Acknowledge to enable FAX BOL.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("hlkMemo")).click();
				Thread.sleep(15000);

				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(5000);

				System.out.println("Memo is working proper.");
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("Memo is not enable.");
				System.out.println("Please ADD Memo from Connect to enable Memo.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("hlkViewCharges")).click();
				Thread.sleep(15000);

				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(5000);

				System.out.println("View Charges is working proper.");
			} catch (Exception e) {
				System.out.println("Please Check Check Box from Connect to view Charges in Net Ship.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("idDocUpload")).click();
				Thread.sleep(20000);

				Driver.findElement(By.id("hlkaddUpload")).click();
				Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
				// DEV
				// Driver.findElement(By.id("file")).sendKeys("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job
				// Upload Doc DEV.xls");
				// Staging
				Driver.findElement(By.id("file")).sendKeys(
						"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job Upload Doc STG.xls");
				Thread.sleep(15000);
				Driver.findElement(By.id("btnUpload")).click();
				Thread.sleep(30000);

				// String Message5 = Driver.findElement(By.id("hrefDocError")).getText();

				// if(Message5.contains("already exists.Your file was saved as"))
				// {
				// Message5 = "*****This File is Already Uploaded Before !*****";
				// System.out.println(Message5);
				// Thread.sleep(5000);
				// }
				// Thread.sleep(5000);

				String Message6 = Driver.findElement(By.id("successid")).getText();

				if (Message6.equals("Import Process Completed !")) {
					Message6 = "*****Import Process is Completed !*****";
					System.out.println(Message6);
					Thread.sleep(5000);
				} else {
					Message6 = Driver.findElement(By.id("errorid")).getText();
					SheetMessage = "*****Import Process is not Completed !*****";
					System.out.println(Message6);
					Thread.sleep(5000);
				}
				Thread.sleep(5000);

				Driver.findElement(By.id("btnSave")).click();
				Thread.sleep(5000);

				System.out.println("Upload link is working proper.");

				if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
					Driver.findElement(By.id("btnCancel")).click();
					Thread.sleep(15000);
				} else {
					Driver.findElement(By.id("btnSave")).click();
					Thread.sleep(15000);

					Driver.findElement(By.id("btnsaveandclose")).click();
					Thread.sleep(15000);
				}
			} catch (Exception e) {
				try {
					Driver.findElement(By.id("iduploadgreen")).click();
					Thread.sleep(20000);

					Driver.findElement(By.id("hlkaddUpload")).click();
					Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
					// DEV
					// Driver.findElement(By.id("file")).sendKeys("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job
					// Upload Doc DEV.xls");
					// Staging
					Driver.findElement(By.id("file")).sendKeys(
							"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job Upload Doc STG.xls");
					Thread.sleep(15000);
					Driver.findElement(By.id("btnUpload")).click();
					Thread.sleep(30000);

					// String Message5 = Driver.findElement(By.id("hrefDocError")).getText();

					// if(Message5.contains("already exists.Your file was saved as"))
					// {
					// Message5 = "*****This File is Already Uploaded Before !*****";
					// System.out.println(Message5);
					// Thread.sleep(5000);
					// }
					// Thread.sleep(5000);

					String Message6 = Driver.findElement(By.id("successid")).getText();

					if (Message6.equals("Import Process Completed !")) {
						Message6 = "*****Import Process is Completed !*****";
						System.out.println(Message6);
						Thread.sleep(5000);
					} else {
						Message6 = Driver.findElement(By.id("errorid")).getText();
						SheetMessage = "*****Import Process is not Completed !*****";
						System.out.println(Message6);
						Thread.sleep(5000);
					}
					Thread.sleep(5000);

					Driver.findElement(By.id("btnSave")).click();
					Thread.sleep(5000);

					System.out.println("Upload link is working proper.");

					if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
						Driver.findElement(By.id("btnCancel")).click();
						Thread.sleep(15000);
					} else {
						Driver.findElement(By.id("btnSave")).click();
						Thread.sleep(15000);

						Driver.findElement(By.id("btnsaveandclose")).click();
						Thread.sleep(15000);
					}
				} catch (Exception f) {
					try {
						Driver.findElement(By.id("hlkUploadDocument")).click();
						Thread.sleep(20000);

						Driver.findElement(By.id("hlkaddUpload")).click();
						Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
						// DEV
						// Driver.findElement(By.id("file")).sendKeys("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job
						// Upload Doc DEV.xls");
						// Staging
						Driver.findElement(By.id("file")).sendKeys(
								"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job Upload Doc STG.xls");
						Thread.sleep(15000);
						Driver.findElement(By.id("btnUpload")).click();
						Thread.sleep(30000);

						// String Message5 = Driver.findElement(By.id("hrefDocError")).getText();

						// if(Message5.contains("already exists.Your file was saved as"))
						// {
						// Message5 = "*****This File is Already Uploaded Before !*****";
						// System.out.println(Message5);
						// Thread.sleep(5000);
						// }
						// Thread.sleep(5000);

						String Message6 = Driver.findElement(By.id("successid")).getText();

						if (Message6.equals("Import Process Completed !")) {
							Message6 = "*****Import Process is Completed !*****";
							System.out.println(Message6);
							Thread.sleep(5000);
						} else {
							Message6 = Driver.findElement(By.id("errorid")).getText();
							SheetMessage = "*****Import Process is not Completed !*****";
							System.out.println(Message6);
							Thread.sleep(5000);
						}
						Thread.sleep(5000);

						Driver.findElement(By.id("btnSave")).click();
						Thread.sleep(5000);

						System.out.println("Upload link is working proper.");

						if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
							Driver.findElement(By.id("btnCancel")).click();
							Thread.sleep(15000);
						} else {
							Driver.findElement(By.id("btnSave")).click();
							Thread.sleep(15000);

							Driver.findElement(By.id("btnsaveandclose")).click();
							Thread.sleep(15000);
						}
					} catch (Exception g) {
						System.out.println("There is no Upload Image display Please check Job manualy.");
					}
				}
			}

			// Click on Watch list
			try {
				Driver.findElement(By.id("watchListBlack")).click();
				Thread.sleep(15000);
				System.out.println("Watch list is working proper.");
			} catch (Exception e) {
				try {
					Driver.findElement(By.id("watchListGreen")).click();
					Thread.sleep(15000);
					System.out.println("Watch list is working proper.");
				} catch (Exception f) {
					System.out.println("There is no Watch List Image display, Because Job is already Delivered.");
				}
			}

			try {
				Driver.findElement(By.id("hlkRefresh")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				System.out.println("There is no job on Active Order from Test Data Excel.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				System.out.println("There is no job on Active Order from Test Data Excel.");
				Thread.sleep(5000);
			}
		}
	}

	@Test
	public void WatchList() throws Exception {
		System.out.println("**********Watch List**********");
		Robot robot = new Robot();
		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipWatchListDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipWatchListSTG.xlsx");

		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("WathcList");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();

		System.out.println("Row Count ====> " + rcount);

		System.out.println("\n********************************************************************************");

		Driver.findElement(By.id("WathcListTab")).click();
		Thread.sleep(20000);

		Select SelectSort0 = new Select(Driver.findElement(By.id("drpcustomer")));
		SelectSort0.selectByIndex(1);
		Thread.sleep(15000);

//		SelectSort0.selectByIndex(2);
//		Thread.sleep(15000);
//		
//		SelectSort0.selectByIndex(3);
//		Thread.sleep(15000);
//		
//		SelectSort0.selectByIndex(4);
//		Thread.sleep(15000);

		SelectSort0.selectByIndex(0);
		Thread.sleep(15000);

		Select SelectSort2 = new Select(Driver.findElement(By.id("drpGrouping")));
		SelectSort2.selectByIndex(1);
		Thread.sleep(15000);

		SelectSort2.selectByIndex(2);
		Thread.sleep(15000);

		SelectSort2.selectByIndex(3);
		Thread.sleep(15000);

		SelectSort2.selectByIndex(4);
		Thread.sleep(15000);

		SelectSort2.selectByIndex(5);
		Thread.sleep(15000);

		SelectSort2.selectByIndex(0);
		Thread.sleep(15000);

		Select SelectSort3 = new Select(Driver.findElement(By.id("drpSorting")));
		SelectSort3.selectByIndex(1);
		Thread.sleep(15000);

		SelectSort3.selectByIndex(2);
		Thread.sleep(15000);

		SelectSort3.selectByIndex(0);
		Thread.sleep(15000);

		for (int i = 1; i <= rcount; i++) {
			System.out.println("\n********************************************************************************");
			System.out.println("\nJob ID ==> " + formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			Thread.sleep(5000);

			String MeJob = "idmemo_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));
			String PrJob = "idprint_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));
			String SeJob = "JobId_" + formatter.formatCellValue(sh0.getRow(i).getCell(0));

			try {
				Driver.findElement(By.id(MeJob)).click();
				Thread.sleep(20000);

				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("There is no Memo Added in Job, Please add Memo first in this Job : "
						+ formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			}

			try {
				Driver.findElement(By.id(PrJob)).click();
				Thread.sleep(5000);

				String winHandleBefore2 = Driver.getWindowHandle();
				for (String winHandle2 : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle2);
				}
				Thread.sleep(15000);

				Driver.close();
				Thread.sleep(15000);
				Driver.switchTo().window(winHandleBefore2);
			} catch (Exception e) {
				System.out.println("Print Label is not able to work on Click, Please check Manualy with this Job : "
						+ formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			}

			List<WebElement> dynamicElement = Driver.findElements(By.id(SeJob));
			if (dynamicElement.size() != 0) {
				Driver.findElement(By.id(SeJob)).click();
				Thread.sleep(20000);

				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(5000);
			} else {
				System.out.println("*********YOUR JOB IS NOT DISPLAY IN LIST.*********");
			}

			System.out.println("\n********************************************************************************");
			System.out.println("\n********************************************************************************");

			Driver.findElement(By.id("txtGlobalSearch")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));
			Driver.findElement(By.id("txtGlobalSearch")).sendKeys(Keys.ENTER);
			Thread.sleep(20000);

			try {
				Driver.findElement(By.id("hlkBillofLading")).click();
				Thread.sleep(5000);

				Driver.findElement(By.id("btemailbol")).click();
				Thread.sleep(5000);

				Driver.findElement(By.id("txtfrom")).clear();
				Driver.findElement(By.id("btnSendBOL")).click();
				Thread.sleep(5000);

				String Message1 = Driver.findElement(By.id("idValidation")).getText();

				if (Message1.contains("Address is required.")) {
					System.out.println("\n*******Validation are display Proper.*******");
					System.out.println("*******" + Message1 + "*******");
					Thread.sleep(5000);
				} else {
					System.out.println("*******Validation are not display Proper.*******");
					System.out.println("*******" + Message1 + "*******");
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("txtfrom")).sendKeys("pdoshi@samyak.com");
				Driver.findElement(By.id("txtto")).sendKeys("puagent@samyakinfo.com,dlagent@samyakinfo.com");
				Driver.findElement(By.id("btnSendBOL")).click();
				Thread.sleep(5000);
				Driver.findElement(By.xpath("/html/body/div[6]/div/div/div[3]/button[1]")).click();
				Thread.sleep(15000);

				String Message2 = Driver.findElement(By.id("saveid")).getText();

				if (Message2.equals("B.O.L # successfully send via Email")) {
					System.out.println("*******Message is display Proper.*******");
					System.out.println("*******" + Message2 + "*******");
				} else {
					System.out.println("*******Validation is not display Proper.*******");
					System.out.println("*******" + Message2 + "*******");
				}
				Thread.sleep(5000);

				Driver.findElement(By.id("btfaxbol")).click();
				Thread.sleep(5000);
				Driver.findElement(By.id("btnSendFAX")).click();
				Thread.sleep(5000);

				String Message3 = Driver.findElement(By.id("errorid")).getText();

				if (Message3.equals("Please select atleast one Fax #s.")) {
					System.out.println("*******Message is display Proper.*******");
					System.out.println("*******" + Message3 + "*******");
				} else {
					System.out.println("*******Validation is not display Proper.*******");
					System.out.println("*******" + Message3 + "*******");
				}
				Thread.sleep(5000);

				// FAX Setup Link
				// Driver.findElement(By.id("hrefFax")).click();
				// Thread.sleep(15000);

				Driver.findElement(By.id("btprint")).click();
				Thread.sleep(5000);

				String winHandleBefore = Driver.getWindowHandle();
				for (String winHandle : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle);
				}
				Thread.sleep(15000);

				Driver.close();
				Thread.sleep(15000);
				Driver.switchTo().window(winHandleBefore);

				Driver.findElement(By.id("btshipdetail")).click();
				Thread.sleep(15000);

				System.out.println("Bill of Lading is working proper.");
			} catch (Exception e) {
				System.out.println("Bill of Lading is not enable.");
				System.out.println("Please Process CSR Acknowledge OR TC Acknowledge to enable Bill of Lading.");
			}

			// Print BOL on Shipment Detail Screen
			try {
				Driver.findElement(By.id("idprintbol")).click();
				Thread.sleep(5000);

				String winHandleBefore1 = Driver.getWindowHandle();
				for (String winHandle1 : Driver.getWindowHandles()) {
					Driver.switchTo().window(winHandle1);
				}
				Thread.sleep(15000);

				Driver.close();
				Thread.sleep(15000);
				Driver.switchTo().window(winHandleBefore1);

				System.out.println("Print BOL is working proper.");
			} catch (Exception e) {
				System.out.println("Print BOL is not enable.");
				System.out.println("Please Process CSR Acknowledge OR TC Acknowledge to enable Print BOL.");
				Thread.sleep(5000);
			}

			// FAX BOL on Shipment Detail Screen
			try {
				Driver.findElement(By.id("hlkFax")).click();
				Thread.sleep(5000);
				Driver.findElement(By.id("btnSendFAXBOL")).click();
				Thread.sleep(15000);

				String Message4 = Driver.findElement(By.id("errorid")).getText();
				Thread.sleep(5000);

				if (Message4.equals("Please select atleast one Fax #s.")) {
					System.out.println("*******Validation Message is display Proper.*******");
					System.out.println("*******" + Message4 + "*******");
				} else {
					System.out.println("*******Validation Message is not display Proper.*******");
					System.out.println("*******" + Message4 + "*******");
				}
				Thread.sleep(5000);

				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(5000);

				System.out.println("FAX BOL is working proper.");
			} catch (Exception e) {
				System.out.println("FAX BOL is not enable.");
				System.out.println("Please Process CSR Acknowledge OR TC Acknowledge to enable FAX BOL.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("hlkMemo")).click();
				Thread.sleep(15000);

				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(5000);

				System.out.println("Memo is working proper.");
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("Memo is not enable.");
				System.out.println("Please ADD Memo from Connect to enable Memo.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("hlkViewCharges")).click();
				Thread.sleep(15000);

				robot.keyPress(KeyEvent.VK_ESCAPE);
				Thread.sleep(5000);

				System.out.println("View Charges is working proper.");
			} catch (Exception e) {
				System.out.println("Please Check Check Box from Connect to view Charges in Net Ship.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("idDocUpload")).click();
				Thread.sleep(20000);

				Driver.findElement(By.id("hlkaddUpload")).click();
				Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
				// DEV
				// Driver.findElement(By.id("file")).sendKeys("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job
				// Upload Doc DEV.xls");
				// Staging
				Driver.findElement(By.id("file")).sendKeys(
						"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job Upload Doc STG.xls");
				Thread.sleep(15000);
				Driver.findElement(By.id("btnUpload")).click();
				Thread.sleep(30000);

				// String Message5 = Driver.findElement(By.id("hrefDocError")).getText();

				// if(Message5.contains("already exists.Your file was saved as"))
				// {
				// Message5 = "*****This File is Already Uploaded Before !*****";
				// System.out.println(Message5);
				// Thread.sleep(5000);
				// }
				// Thread.sleep(5000);

				String Message6 = Driver.findElement(By.id("successid")).getText();

				if (Message6.equals("Import Process Completed !")) {
					Message6 = "*****Import Process is Completed !*****";
					System.out.println(Message6);
					Thread.sleep(5000);
				} else {
					Message6 = Driver.findElement(By.id("errorid")).getText();
					SheetMessage = "*****Import Process is not Completed !*****";
					System.out.println(Message6);
					Thread.sleep(5000);
				}
				Thread.sleep(5000);

				Driver.findElement(By.id("btnSave")).click();
				Thread.sleep(5000);

				System.out.println("Upload link is working proper.");

				if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
					Driver.findElement(By.id("btnCancel")).click();
					Thread.sleep(15000);
				} else {
					Driver.findElement(By.id("btnSave")).click();
					Thread.sleep(15000);

					Driver.findElement(By.id("btnsaveandclose")).click();
					Thread.sleep(15000);
				}
			} catch (Exception e) {
				try {
					Driver.findElement(By.id("iduploadgreen")).click();
					Thread.sleep(20000);

					Driver.findElement(By.id("hlkaddUpload")).click();
					Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
					// DEV
					// Driver.findElement(By.id("file")).sendKeys("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job
					// Upload Doc DEV.xls");
					// Staging
					Driver.findElement(By.id("file")).sendKeys(
							"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job Upload Doc STG.xls");
					Thread.sleep(15000);
					Driver.findElement(By.id("btnUpload")).click();
					Thread.sleep(30000);

					// String Message5 = Driver.findElement(By.id("hrefDocError")).getText();

					// if(Message5.contains("already exists.Your file was saved as"))
					// {
					// Message5 = "*****This File is Already Uploaded Before !*****";
					// System.out.println(Message5);
					// Thread.sleep(5000);
					// }
					// Thread.sleep(5000);

					String Message6 = Driver.findElement(By.id("successid")).getText();

					if (Message6.equals("Import Process Completed !")) {
						Message6 = "*****Import Process is Completed !*****";
						System.out.println(Message6);
						Thread.sleep(5000);
					} else {
						Message6 = Driver.findElement(By.id("errorid")).getText();
						SheetMessage = "*****Import Process is not Completed !*****";
						System.out.println(Message6);
						Thread.sleep(5000);
					}
					Thread.sleep(5000);

					Driver.findElement(By.id("btnSave")).click();
					Thread.sleep(5000);

					System.out.println("Upload link is working proper.");

					if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
						Driver.findElement(By.id("btnCancel")).click();
						Thread.sleep(15000);
					} else {
						Driver.findElement(By.id("btnSave")).click();
						Thread.sleep(15000);

						Driver.findElement(By.id("btnsaveandclose")).click();
						Thread.sleep(15000);
					}
				} catch (Exception f) {
					try {
						Driver.findElement(By.id("hlkUploadDocument")).click();
						Thread.sleep(20000);

						Driver.findElement(By.id("hlkaddUpload")).click();
						Driver.findElement(By.id("txtDocName")).sendKeys("AUTOPdoshiDocument");
						// DEV
						// Driver.findElement(By.id("file")).sendKeys("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job
						// Upload Doc DEV.xls");
						// Staging
						Driver.findElement(By.id("file")).sendKeys(
								"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Job Upload Doc STG.xls");
						Thread.sleep(15000);
						Driver.findElement(By.id("btnUpload")).click();
						Thread.sleep(30000);

						// String Message5 = Driver.findElement(By.id("hrefDocError")).getText();

						// if(Message5.contains("already exists.Your file was saved as"))
						// {
						// Message5 = "*****This File is Already Uploaded Before !*****";
						// System.out.println(Message5);
						// Thread.sleep(5000);
						// }
						// Thread.sleep(5000);

						String Message6 = Driver.findElement(By.id("successid")).getText();

						if (Message6.equals("Import Process Completed !")) {
							Message6 = "*****Import Process is Completed !*****";
							System.out.println(Message6);
							Thread.sleep(5000);
						} else {
							Message6 = Driver.findElement(By.id("errorid")).getText();
							SheetMessage = "*****Import Process is not Completed !*****";
							System.out.println(Message6);
							Thread.sleep(5000);
						}
						Thread.sleep(5000);

						Driver.findElement(By.id("btnSave")).click();
						Thread.sleep(5000);

						System.out.println("Upload link is working proper.");

						if (Driver.findElement(By.id("hrefDocError")).isDisplayed()) {
							Driver.findElement(By.id("btnCancel")).click();
							Thread.sleep(15000);
						} else {
							Driver.findElement(By.id("btnSave")).click();
							Thread.sleep(15000);

							Driver.findElement(By.id("btnsaveandclose")).click();
							Thread.sleep(15000);
						}
					} catch (Exception g) {
						System.out.println("There is no Upload Image display Please check Job manualy.");
					}
				}
			}

			try {
				Driver.findElement(By.id("hlkRefresh")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				System.out.println("There is no job on Active Order from Test Data Excel.");
				Thread.sleep(15000);
			}

			try {
				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				System.out.println("There is no job on Active Order from Test Data Excel.");
				Thread.sleep(15000);
			}
		}
	}

	@Test
	public void BasicSearch0() throws Exception {
		System.out.println("**********Basic Search**********");
		Robot robot = new Robot();
		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipBasicSearchDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipBasicSearchSTG.xlsx");
		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("GetQuote");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();
		System.out.println("Row Count ====> " + rcount);

		System.out
				.println("******************************************************************************************");
		Driver.findElement(By.id("FindOptionTab")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("btnContinue")).click();
		Thread.sleep(20000);

		String Message1 = Driver.findElement(By.id("errorid")).getText();

		if (Message1.contains("Zipcode/Airport.")) {
			System.out.println(
					"******************************************************************************************");
			System.out.println("Press Get Quote button without enter Zip Codes.");
			System.out.println(
					"******************************************************************************************");
			System.out.println(Message1);
			System.out.println(
					"******************************************************************************************");
			Thread.sleep(5000);
		}

		Driver.findElement(By.id("txtPUZipCode")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(0)));
		Thread.sleep(5000);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtDLZipCode")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtDLZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(1)));
		Thread.sleep(5000);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);

		Driver.findElement(By.id("btnContinue")).click();
		Thread.sleep(20000);

		String Message2 = Driver.findElement(By.id("errorid")).getText();

		if (Message2.contains("Zipcode/Airport.")) {
			System.out.println(
					"******************************************************************************************");
			System.out.println("Press Get Quote button with invalid Zip Codes.");
			System.out.println(
					"******************************************************************************************");
			System.out.println(Message2);
			System.out.println(
					"******************************************************************************************");
			Thread.sleep(5000);
		}
		Thread.sleep(5000);
		System.out
				.println("******************************************************************************************");

		for (int i = 2; i <= rcount; i++) {
			System.out.println(
					"******************************************************************************************");

			Driver.findElement(By.id("txtPUZipCode")).clear();
			Thread.sleep(5000);
			Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			Thread.sleep(5000);
			robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(5000);

			Driver.findElement(By.id("txtDLZipCode")).clear();
			Thread.sleep(5000);
			Driver.findElement(By.id("txtDLZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));
			Thread.sleep(5000);
			robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(5000);

			Driver.findElement(By.id("chkPickupItemsAll")).click();
			Thread.sleep(5000);
			Driver.findElement(By.id("hlkdeletePackage")).click();
			Thread.sleep(5000);
			Random rand = new Random();
			int k = 0;
			for (int j = 1; j <= 5; j++) {
				Driver.findElement(By.id("hlkaddPackage")).click();
				Thread.sleep(5000);

				String Qu, Pt, We, Le, Wi, He;
				Qu = "txtQuantity_" + k;
				Pt = "drpPackageType_" + k;
				We = "txtWeight_" + k;
				Le = "txtDimLength_" + k;
				Wi = "txtDimWidth_" + k;
				He = "txtDimHeight_" + k;

				Select PaTy = new Select(Driver.findElement(By.id(Pt)));
				PaTy.selectByIndex(k);
				Thread.sleep(5000);
				int num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				String num2 = Integer.toString(num1);
				Driver.findElement(By.id(Qu)).clear();
				Driver.findElement(By.id(Qu)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(We)).clear();
				Driver.findElement(By.id(We)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(Le)).clear();
				Driver.findElement(By.id(Le)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(Wi)).clear();
				Driver.findElement(By.id(Wi)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(He)).clear();
				Driver.findElement(By.id(He)).sendKeys(num2);
				Thread.sleep(5000);

				k++;
			}

			Driver.findElement(By.id("btnContinue")).click();
			Thread.sleep(20000);

			getscreenshot("GetQuote_" + i);

			String Message3 = Driver.findElement(By.id("errorid")).getText();

			if (Message3.contains("Zipcode/Airport.")) {
				System.out.println(
						"******************************************************************************************");
				System.out.println("Press Get Quote button with invalid Zip Codes.");
				System.out.println(
						"******************************************************************************************");
				System.out.println(Message3);
				System.out.println(
						"******************************************************************************************");
				Thread.sleep(5000);
			}

			try {
				String Data1 = Driver.findElement(By.id("tblFOGroundOptions")).getText();
				System.out.println(Data1);
				if (Data1.contains("Local delivery")) {
					System.out.println(
							"******************************************************************************************");
					System.out.println(
							"BOTH selected Zip Codes have Mileage less than 250, So Both are for Ground Services.");
					System.out.println(
							"******************************************************************************************");
					Thread.sleep(5000);
				} else if (Data1.contains(" AIR ")) {
					System.out.println(
							"******************************************************************************************");
					System.out.println(
							"BOTH selected Zip Codes have Mileage more than 250, So Both are for AIR Services.");
					System.out.println(
							"******************************************************************************************");
					Thread.sleep(5000);
				} else {
					System.out.println(
							"******************************************************************************************");
					System.out.println("Need to Check manualy ====> " + Data1);
					System.out.println(
							"******************************************************************************************");
					Thread.sleep(5000);
				}
			} catch (Exception e) {
				System.out.println("There is no data display for Rate May be Zip Codes are invalid.");
			}

			Thread.sleep(5000);
			System.out.println(
					"******************************************************************************************");
		}
	}

	@Test
	public void BasicSearch1() throws Exception {
		System.out.println("**********Basic Search**********");
		Robot robot = new Robot();
		Driver.findElement(By.id("imgLogo")).click();
		Thread.sleep(5000);

		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipBasicSearchDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipBasicSearchSTG.xlsx");
		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("GetQuote");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();
		System.out.println("Row Count ====> " + rcount);

		System.out
				.println("******************************************************************************************");
		Driver.findElement(By.id("FindOptionTab")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("advanceTab")).click();
		Thread.sleep(20000);

		try {
			Driver.findElement(By.id("btnContinue")).click();
			Thread.sleep(20000);
		} catch (Exception e) {
			JavascriptExecutor jse = (JavascriptExecutor) Driver;
			jse.executeScript("window.scrollBy(0,350)");
			Driver.findElement(By.id("btnContinue")).click();
			Thread.sleep(20000);
		}

		String Message1 = Driver.findElement(By.id("errorid")).getText();

		if (Message1.contains("Zipcode")) {
			System.out.println(
					"******************************************************************************************");
			System.out.println("Press Advance Search button without enter Zip Codes.");
			System.out.println(
					"******************************************************************************************");
			System.out.println(Message1);
			System.out.println(
					"******************************************************************************************");
			Thread.sleep(5000);
		}

		Driver.findElement(By.id("txtPUZipCode")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(0)));
		Thread.sleep(15000);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtDLZipCode")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtDLZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(1)));
		Thread.sleep(5000);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(15000);

		try {
			Driver.findElement(By.id("btnContinue")).click();
			Thread.sleep(20000);
		} catch (Exception e) {
			JavascriptExecutor jse = (JavascriptExecutor) Driver;
			jse.executeScript("window.scrollBy(0,350)");
			Driver.findElement(By.id("btnContinue")).click();
			Thread.sleep(20000);
		}

		String Message2 = Driver.findElement(By.id("errorid")).getText();

		if (Message2.contains("Zipcode")) {
			System.out.println(
					"******************************************************************************************");
			System.out.println("Press Advance Search button with invalid Zip Codes.");
			System.out.println(
					"******************************************************************************************");
			System.out.println(Message2);
			System.out.println(
					"******************************************************************************************");
			Thread.sleep(5000);
		}
		Thread.sleep(5000);
		System.out
				.println("******************************************************************************************");

		for (int i = 2; i <= rcount; i++) {
			System.out.println(
					"******************************************************************************************");

			Driver.findElement(By.id("txtPUZipCode")).clear();
			Thread.sleep(5000);
			Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			Thread.sleep(5000);
			robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(15000);

			Driver.findElement(By.id("txtDLZipCode")).clear();
			Thread.sleep(5000);
			Driver.findElement(By.id("txtDLZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));
			Thread.sleep(5000);
			robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(15000);

			try {
				Driver.findElement(By.id("chkPickupItemsAll")).click();
				Thread.sleep(5000);
				Driver.findElement(By.id("hlkdeletePackage")).click();
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("There is no default Packge on Basic Search Advance Screen.");
			}
			Random rand = new Random();
			int k = 0;
			for (int j = 1; j <= 5; j++) {
				Driver.findElement(By.id("hlkaddPackage")).click();
				Thread.sleep(5000);

				String Qu, Pt, We, Le, Wi, He;
				Qu = "txtQuantity_" + k;
				Pt = "drpPackageType_" + k;
				We = "txtWeight_" + k;
				Le = "txtDimLength_" + k;
				Wi = "txtDimWidth_" + k;
				He = "txtDimHeight_" + k;

				Select PaTy = new Select(Driver.findElement(By.id(Pt)));
				PaTy.selectByIndex(k);
				Thread.sleep(5000);
				int num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				String num2 = Integer.toString(num1);
				Driver.findElement(By.id(Qu)).clear();
				Driver.findElement(By.id(Qu)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(We)).clear();
				Driver.findElement(By.id(We)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(Le)).clear();
				Driver.findElement(By.id(Le)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(Wi)).clear();
				Driver.findElement(By.id(Wi)).sendKeys(num2);
				Thread.sleep(5000);
				num1 = rand.nextInt(20);
				if (num1 == 0) {
					num1 = num1 + 2;
				}
				num2 = Integer.toString(num1);
				Driver.findElement(By.id(He)).clear();
				Driver.findElement(By.id(He)).sendKeys(num2);
				Thread.sleep(5000);

				k++;
			}
			int ss = 1;

			try {
				Driver.findElement(By.id("btnContinue")).click();
				Thread.sleep(20000);
			} catch (Exception e) {
				JavascriptExecutor jse = (JavascriptExecutor) Driver;
				jse.executeScript("window.scrollBy(0,350)");
				Driver.findElement(By.id("btnContinue")).click();
				Thread.sleep(20000);
			}

			getscreenshot("GetQuote_" + ss);

			String Message3 = Driver.findElement(By.id("errorid")).getText();

			if (Message3.contains("Zipcode/Airport.")) {
				System.out.println(
						"******************************************************************************************");
				System.out.println("Press Get Quote button with invalid Zip Codes.");
				System.out.println(
						"******************************************************************************************");
				System.out.println(Message3);
				System.out.println(
						"******************************************************************************************");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("btnShipContinue")).click();
				Thread.sleep(20000);
			} catch (Exception e) {
				JavascriptExecutor jse = (JavascriptExecutor) Driver;
				jse.executeScript("window.scrollBy(0,350)");
				Driver.findElement(By.id("btnShipContinue")).click();
				Thread.sleep(20000);
			}

			ss++;
			getscreenshot("GetQuote_" + ss);
			Thread.sleep(5000);
			try {
				String Data1 = Driver.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div/div[3]/div/div/div[2]/div[3]/div[2]/div/div/div[2]/div[1]/table/tbody/tr/td[5]"))
						.getText();
				System.out.println(Data1);
				if (Data1.contains("Local delivery")) {
					System.out.println(
							"******************************************************************************************");
					System.out.println(
							"BOTH selected Zip Codes have Mileage less than 250, So Both are for Ground Services.");
					System.out.println(
							"******************************************************************************************");
					Thread.sleep(5000);
				} else if (Data1.contains(" AIR ")) {
					System.out.println(
							"******************************************************************************************");
					System.out.println(
							"BOTH selected Zip Codes have Mileage more than 250, So Both are for AIR Services.");
					System.out.println(
							"******************************************************************************************");
					Thread.sleep(5000);
				} else {
					System.out.println(
							"******************************************************************************************");
					System.out.println("Need to Check manualy ====> " + Data1);
					System.out.println(
							"******************************************************************************************");
					Thread.sleep(5000);
				}
			} catch (Exception e) {
				System.out.println("There is no data display for Rate May be Zip Codes are invalid.");
			}
			Thread.sleep(5000);
			try {
				Driver.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div/div[3]/div/div/div[2]/div[3]/div[2]/div/div/div[2]/div[2]/div/span"))
						.getText();
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("There is no Rate display on Screen.");
				Thread.sleep(5000);
			}

			try {
				Driver.findElement(By.id("hlkSerachAgain")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				JavascriptExecutor jse = (JavascriptExecutor) Driver;
				jse.executeScript("window.scrollBy(0,-350)");
				Driver.findElement(By.id("hlkSerachAgain")).click();
				Thread.sleep(15000);
			}
			System.out.println(
					"******************************************************************************************");
		}
	}

	@Test
	public void BasicSearch2() throws Exception {
		System.out.println("**********Basic Search**********");
		Robot robot = new Robot();
		Driver.findElement(By.id("imgLogo")).click();
		Thread.sleep(5000);

		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipBasicSearchDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipBasicSearchSTG.xlsx");
		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("GetQuote");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();
		System.out.println("Row Count ====> " + rcount);

		System.out
				.println("******************************************************************************************");
		Driver.findElement(By.id("FindOptionTab")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("RateCalc")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnCalculator")).click();
		Thread.sleep(20000);

		String Message1 = Driver.findElement(By.id("errorid")).getText();

		if (Message1.contains("Required")) {
			System.out.println(
					"******************************************************************************************");
			System.out.println("Press Calculate button without enter Data.");
			System.out.println(
					"******************************************************************************************");
			System.out.println(Message1);
			System.out.println(
					"******************************************************************************************");
			Thread.sleep(5000);
		}

		Driver.findElement(By.id("txtpieces")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtpieces")).sendKeys("5");
		Thread.sleep(5000);

		Driver.findElement(By.id("txtweight")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtweight")).sendKeys("55");
		Thread.sleep(5000);

		Driver.findElement(By.id("txtPUZipCode")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(0)));
		Thread.sleep(5000);

		String Message2 = Driver.findElement(By.id("errorid")).getText();

		if (Message2.contains("Invalid")) {
			System.out.println(
					"******************************************************************************************");
			System.out.println("Press Rate Calculator button with invalid Zip Codes.");
			System.out.println(
					"******************************************************************************************");
			System.out.println(Message2);
			System.out.println(
					"******************************************************************************************");
			Thread.sleep(5000);
		}
		Driver.findElement(By.id("txtPUZipCode")).clear();
		Thread.sleep(5000);

		Driver.findElement(By.id("txtDLZipCode")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtDLZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(1)));
		Thread.sleep(5000);

		String Message3 = Driver.findElement(By.id("errorid")).getText();

		if (Message3.contains("Invalid")) {
			System.out.println(
					"******************************************************************************************");
			System.out.println("Press Rate Calculator button with invalid Zip Codes.");
			System.out.println(
					"******************************************************************************************");
			System.out.println(Message3);
			System.out.println(
					"******************************************************************************************");
			Thread.sleep(5000);
		}
		Thread.sleep(5000);
		System.out
				.println("******************************************************************************************");

		for (int i = 2; i <= rcount; i++) {
			System.out.println(
					"******************************************************************************************");

			Driver.findElement(By.id("txtPUZipCode")).clear();
			Thread.sleep(5000);
			Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			Thread.sleep(5000);
			robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(5000);

			Driver.findElement(By.id("txtDLZipCode")).clear();
			Thread.sleep(5000);
			Driver.findElement(By.id("txtDLZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));
			Thread.sleep(5000);
			robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(5000);

			if (i == 2) {
				String ServiceName = "Rush Drive - LOC";
				Select dropdown1 = new Select(Driver.findElement(By.id("cmbservice")));
				dropdown1.selectByVisibleText(ServiceName);
				Thread.sleep(5000);
				System.out.println("LOC Service is display proper.");
			} else {
				String ServiceName = "Next Flight Out - SD";
				Select dropdown1 = new Select(Driver.findElement(By.id("cmbservice")));
				dropdown1.selectByVisibleText(ServiceName);
				Thread.sleep(5000);
				System.out.println("SD Service is display proper.");
			}

			Driver.findElement(By.id("btnCalculator")).click();
			Thread.sleep(20000);

			if (Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[3]/div/div/table/thead[2]/tr/th[4]"))
					.isDisplayed() == true) {
				Thread.sleep(5000);
				System.out.println(Driver.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[3]/div/div/table/thead[2]/tr/th[4]"))
						.getText());
				Thread.sleep(5000);
			} else {
				System.out.println("Rate are not display.");
				Thread.sleep(5000);
			}

			getscreenshot("GetQuote_" + i);

			Thread.sleep(5000);
			System.out.println(
					"******************************************************************************************");
		}
	}

	@Test
	public void ShipPackage() throws Exception {
		Robot robot = new Robot();
		// click on Ship Package
		Driver.findElement(By.id("ShipPackageTab")).click();
		Thread.sleep(15000);

		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NSJobDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NSJobSTG.xlsx");

		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("Sheet1");
		int rcount = sh0.getLastRowNum();

		System.out.println("Row Count ====> " + rcount);

		for (int i = 1; i <= rcount; i++) {
			DataFormatter formatter = new DataFormatter();
			Select dropdown = new Select(Driver.findElement(By.id("drpClient")));
			dropdown.selectByVisibleText(CustomerNameNSPL);
			Thread.sleep(15000);

			Driver.findElement(By.id("txtorderplace")).clear();
			Driver.findElement(By.id("txtorderplace")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(0)));

			Driver.findElement(By.id("txtPickUpPhone")).clear();
			Driver.findElement(By.id("txtPickUpPhone")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));

			System.out.println(Driver.findElement(By.id("btnPuDelInfo")).getText());
			Driver.findElement(By.id("btnPuDelInfo")).click();
			Thread.sleep(15000);

			Driver.findElement(By.id("txtPUCompanyName")).clear();
			Driver.findElement(By.id("txtPUCompanyName")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(2)));
			Driver.findElement(By.id("txtaddressline")).clear();
			Driver.findElement(By.id("txtaddressline")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(3)));
			Driver.findElement(By.id("txtdeptsuite")).clear();
			Driver.findElement(By.id("txtdeptsuite")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(4)));
			Driver.findElement(By.id("txtPUZipCode")).clear();
			Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(5)));
			Thread.sleep(15000);
			Driver.findElement(By.id("txtPUZipCode")).sendKeys(Keys.ENTER);
			Driver.findElement(By.id("txtpertosee")).clear();
			Driver.findElement(By.id("txtpertosee")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(6)));
			Driver.findElement(By.id("txtphone")).clear();
			Driver.findElement(By.id("txtphone")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(7)));
			Driver.findElement(By.id("txtnote")).clear();
			Driver.findElement(By.id("txtnote")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(9)));
			Driver.findElement(By.id("txtEmail")).clear();
			Driver.findElement(By.id("txtEmail")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(8)));

			Driver.findElement(By.id("txtDLCompanyName")).clear();
			Driver.findElement(By.id("txtDLCompanyName"))
					.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(10)));
			Driver.findElement(By.id("txtaddresslinedel")).clear();
			Driver.findElement(By.id("txtaddresslinedel"))
					.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(11)));
			Driver.findElement(By.id("txtdeptsuitedel")).clear();
			Driver.findElement(By.id("txtdeptsuitedel")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(12)));
			Driver.findElement(By.id("txtDLZipCode")).clear();
			Driver.findElement(By.id("txtDLZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(13)));
			Thread.sleep(15000);
			Driver.findElement(By.id("txtDLZipCode")).sendKeys(Keys.ENTER);
			Driver.findElement(By.id("txtpertoseedel")).clear();
			Driver.findElement(By.id("txtpertoseedel")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(14)));
			Driver.findElement(By.id("txtphonedel")).clear();
			Driver.findElement(By.id("txtphonedel")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(15)));
			Driver.findElement(By.id("txtinstructiondel")).clear();
			Driver.findElement(By.id("txtinstructiondel"))
					.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(17)));
			Driver.findElement(By.id("txtEmaildel")).clear();
			Driver.findElement(By.id("txtEmaildel")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(16)));

			String ServiceName = "Rush Drive - LOC";

			Select dropdown1 = new Select(Driver.findElement(By.id("cmbdefualtservice")));
			dropdown1.selectByVisibleText(ServiceName);

			Thread.sleep(15000);
			Driver.findElement(By.id("txtpertoseedel")).sendKeys(Keys.PAGE_DOWN);

			try {
				Driver.findElement(By.id("btnPlaceOrder")).click();
				Thread.sleep(15000);
			} catch (Exception e) {
				robot.keyPress(KeyEvent.VK_CONTROL + KeyEvent.VK_END);
				Thread.sleep(5000);
				Driver.findElement(By.id("btnPlaceOrder")).click();
				Thread.sleep(15000);
			}

			String JobCreatedMess = Driver.findElement(By.id("confirmorderpopup")).getText();
			System.out.println(JobCreatedMess);

			String[] list = JobCreatedMess.split(" ");
			System.out.println("Pickup# :- " + list[15]);

			Driver.findElement(By.id("idordercancelbtn")).click();
			Thread.sleep(15000);
		}
		Thread.sleep(5000);
	}

	@Test
	public void AllReports() throws Exception {
		System.out.println("********************ALL Reports********************");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);

		System.out.println("**********Activity Summary**********");
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Activity Summary")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Activity Summary");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Order Details**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Order Details")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(5000);

		getscreenshot("Order Details1");

		String Message1 = Driver.findElement(By.id("idValidation")).getText();

		if (Message1.contains("Required")) {
			System.out.println(Message1);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Select dropdown1 = new Select(Driver.findElement(By.id("drpClient")));
		dropdown1.selectByIndex(1);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Order Details2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Month-to-Month**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Month-to-Month")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Month-to-Month");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Root Cause**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Root Cause")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Root Cause");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Pull Time Performance**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Pull Time Performance")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Pull Time Performance1");

		String Message2 = Driver.findElement(By.id("idValidation")).getText();

		if (Message2.contains("Please")) {
			System.out.println(Message2);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlService")).click();
		Thread.sleep(5000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div[2]/div[2]/div/div/ul/li[1]/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Pull Time Performance2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Receipts**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Receipts")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Receipts");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Out bounds**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Outbounds")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Outbounds");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Transactions**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Transactions")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Transactions1");

		String Message3 = Driver.findElement(By.id("idValidation")).getText();

		if (Message3.contains("Please")) {
			System.out.println(Message3);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div[1]/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Transactions2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Inventory Audit**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Inventory Audit")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Inventory Audit1");

		String Message4 = Driver.findElement(By.id("idValidation")).getText();

		if (Message4.contains("Please")) {
			System.out.println(Message4);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div[1]/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Inventory Audit2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Inventory on Hand**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Inventory on Hand")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Inventory on Hand1");

		String Message5 = Driver.findElement(By.id("idValidation")).getText();

		if (Message5.contains("Please")) {
			System.out.println(Message5);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Inventory on Hand2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Replenish Parts**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Replenish Parts")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Replenish Parts1");

		String Message6 = Driver.findElement(By.id("idValidation")).getText();

		if (Message6.contains("Please")) {
			System.out.println(Message6);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Replenish Parts2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Quarantine**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Quarantine")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Quarantine1");

		String Message7 = Driver.findElement(By.id("idValidation")).getText();

		if (Message7.contains("Please")) {
			System.out.println(Message7);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("Quarantine2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********By Location**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("By Location")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("By Location");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********By Model/Part#**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		try {
			Driver.findElement(By.linkText("By Model/Part#")).click();
			Thread.sleep(15000);
		} catch (Exception e) {
			JavascriptExecutor jse = (JavascriptExecutor) Driver;
			jse.executeScript("window.scrollBy(0,350)");
			Driver.findElement(By.linkText("By Model/Part#")).click();
			Thread.sleep(15000);
		}

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("By Model-Part#1");

		String Message11 = Driver.findElement(By.id("idValidation")).getText();

		if (Message11.contains("Please")) {
			System.out.println(Message11);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("By Model-Part#2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********In Transit**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		try {
			Driver.findElement(By.linkText("In Transit")).click();
			Thread.sleep(15000);
		} catch (Exception e) {
			JavascriptExecutor jse = (JavascriptExecutor) Driver;
			jse.executeScript("window.scrollBy(0,350)");
			Driver.findElement(By.linkText("In Transit")).click();
			Thread.sleep(15000);
		}

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("In Transit1");

		String Message8 = Driver.findElement(By.id("idValidation")).getText();

		if (Message8.contains("Please")) {
			System.out.println(Message8);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("In Transit2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********FSL Address**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		try {
			Driver.findElement(By.linkText("FSL Address")).click();
			Thread.sleep(15000);
		} catch (Exception e) {
			JavascriptExecutor jse = (JavascriptExecutor) Driver;
			jse.executeScript("window.scrollBy(0,350)");
			Driver.findElement(By.linkText("FSL Address")).click();
			Thread.sleep(15000);
		}

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("FSL Address1");

		String Message9 = Driver.findElement(By.id("idValidation")).getText();

		if (Message9.contains("Please")) {
			System.out.println(Message9);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div/div[2]/div/form/div[1]/div/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnView")).click();
		Thread.sleep(20000);

		getscreenshot("FSL Address2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Report Builder**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Report Builder")).click();
		Thread.sleep(15000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div[2]/div[2]/div[2]/div/div/table/tbody/tr[1]/td[1]/input"))
				.click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnGenReport")).click();
		Thread.sleep(25000);

		getscreenshot("Report Builder");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Export Inventory**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Export Inventory")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnGetFSL")).click();
		Thread.sleep(20000);

		getscreenshot("Export Inventory1");

		String Message10 = Driver.findElement(By.id("idValidation")).getText();

		if (Message10.contains("Please")) {
			System.out.println(Message10);
		} else {
			System.out.println("There is no Error Message display.");
		}

		Driver.findElement(By.id("ddlClient")).click();
		Thread.sleep(5000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[3]/div[1]/div/div/div[1]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnGetFSL")).click();
		Thread.sleep(20000);

		getscreenshot("Export Inventory2");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}

		System.out.println("**********Shipping Data Download**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgreport")).click();
		Thread.sleep(5000);
		Driver.findElement(By.linkText("Shipping Data Download")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnDownload")).click();
		Thread.sleep(20000);

		getscreenshot("Shipping Data Download");

		try {
			if (Driver.findElement(By.id("myIframe")).isDisplayed() == true) {
				System.out.println("Report is display proper with values......");
			} else {
				System.out.println("Report is not display proper with values......");
				System.out.println("Check Screenshot......");
			}
		} catch (Exception e) {
			System.out.println("Report Response is not display......");
			System.out.println("Check Screenshot......");
		}
	}

	@Test
	public void AddressBook() throws Exception {
		System.out.println("**********Address Book**********");
		Robot robot = new Robot();
		Random rand = new Random();
		int a = 0, b = 0, a1 = 0, b1 = 0, c, d;
		Driver.findElement(By.id("imgaddressbook")).click();
		Thread.sleep(35000);

		Driver.findElement(By.id("imgNew")).click();
		Thread.sleep(15000);

		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("AddressBook1");
			Thread.sleep(5000);
			System.out.println("Address Book FIRST validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book FIRST validation is not display proper.");
			getscreenshot("AddressBook2");
			Thread.sleep(5000);
		}

		Select dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameNSPL);
		Thread.sleep(20000);
		if (Driver
				.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
				.isDisplayed() == true) {
			System.out.println("Address Count is display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			String[] List1 = Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText().split(" ");
			a = Integer.parseInt(List1[3]);
			System.out.println("First Count ==>" + a);
			getscreenshot("AddressBook3");
			Thread.sleep(5000);
		} else {
			System.out.println("Address Count is not display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			getscreenshot("AddressBook4");
			Thread.sleep(5000);
		}

		dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameSPL);
		Thread.sleep(20000);
		if (Driver
				.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
				.isDisplayed() == true) {
			System.out.println("Address Count is display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			String[] List1 = Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText().split(" ");
			b = Integer.parseInt(List1[3]);
			System.out.println("Second Count ==>" + b);
			getscreenshot("AddressBook5");
			Thread.sleep(5000);
		} else {
			System.out.println("Address Count is not display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			getscreenshot("AddressBook6");
			Thread.sleep(5000);
		}

		dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameNSPL);
		Thread.sleep(5000);
		Driver.findElement(By.id("imgNew")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("AddressBook7");
			Thread.sleep(5000);
			System.out.println("Address Book SECOND validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book SECOND validation is not display proper.");
			getscreenshot("AddressBook8");
			Thread.sleep(5000);
		}
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		int num1 = rand.nextInt(200);
		Driver.findElement(By.id("txtbusinessname")).sendKeys("PDOSHIBN" + num1);
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("AddressBook9");
			Thread.sleep(5000);
			System.out.println("Address Book THIRD validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book THIRD validation is not display proper.");
			getscreenshot("AddressBook10");
			Thread.sleep(5000);
		}
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		Driver.findElement(By.id("txtZipCode")).sendKeys("21225");
		Thread.sleep(5000);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("AddressBook11");
			Thread.sleep(5000);
			System.out.println("Address Book FOURTH validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book FOURTH validation is not display proper.");
			getscreenshot("AddressBook12");
			Thread.sleep(5000);
		}
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		int num2 = rand.nextInt(20000);
		Driver.findElement(By.id("txtaddressline")).sendKeys(num2 + ", EMMA WATSON APP.");
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("success")).isDisplayed() == true) {
			getscreenshot("AddressBook13");
			Thread.sleep(5000);
			System.out.println("Address Book FIFTH validation is display proper.");
			String Message1 = Driver.findElement(By.id("success")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book FIFTH validation is not display proper.");
			try {
				String Message1 = Driver.findElement(By.id("success")).getText();
				System.out.println(Message1);
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("There is no SUCCESS Message display.");
			}
			try {
				String Message1 = Driver.findElement(By.id("idValidation")).getText();
				System.out.println(Message1);
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("There is no FAILED Message display.");
			}
			getscreenshot("AddressBook14");
			Thread.sleep(5000);
		}

		dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameSPL);
		Thread.sleep(5000);
		Driver.findElement(By.id("imgNew")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("AddressBook15");
			Thread.sleep(5000);
			System.out.println("Address Book SECOND validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book SECOND validation is not display proper.");
			getscreenshot("AddressBook16");
			Thread.sleep(5000);
		}
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		int num3 = rand.nextInt(200);
		Driver.findElement(By.id("txtbusinessname")).sendKeys("PDOSHIBN" + num3);
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("AddressBook17");
			Thread.sleep(5000);
			System.out.println("Address Book THIRD validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book THIRD validation is not display proper.");
			getscreenshot("AddressBook18");
			Thread.sleep(5000);
		}
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		Driver.findElement(By.id("txtZipCode")).sendKeys("21225");
		Thread.sleep(5000);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("AddressBook19");
			Thread.sleep(5000);
			System.out.println("Address Book FOURTH validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book FOURTH validation is not display proper.");
			getscreenshot("AddressBook20");
			Thread.sleep(5000);
		}
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		int num4 = rand.nextInt(20000);
		Driver.findElement(By.id("txtaddressline")).sendKeys(num4 + ", EMMA WATSON APP.");
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		if (Driver.findElement(By.id("success")).isDisplayed() == true) {
			getscreenshot("AddressBook21");
			Thread.sleep(5000);
			System.out.println("Address Book FIFTH validation is display proper.");
			String Message1 = Driver.findElement(By.id("success")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Address Book FIFTH validation is not display proper.");
			try {
				String Message1 = Driver.findElement(By.id("success")).getText();
				System.out.println(Message1);
			} catch (Exception e) {
				System.out.println("There is no SUCCESS Message display.");
			}
			try {
				String Message1 = Driver.findElement(By.id("idValidation")).getText();
				System.out.println(Message1);
			} catch (Exception e) {
				System.out.println("There is no FAILED Message display.");
			}
			getscreenshot("AddressBook22");
			Thread.sleep(5000);
		}

		dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameNSPL);
		Thread.sleep(20000);
		if (Driver
				.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
				.isDisplayed() == true) {
			System.out.println("Address Count is display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			String[] List1 = Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText().split(" ");
			a1 = Integer.parseInt(List1[3]);
			System.out.println("First Count ==>" + a1);
			getscreenshot("AddressBook23");
			Thread.sleep(5000);
		} else {
			System.out.println("Address Count is not display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			getscreenshot("AddressBook24");
			Thread.sleep(5000);
		}

		dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameSPL);
		Thread.sleep(20000);
		if (Driver
				.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
				.isDisplayed() == true) {
			System.out.println("Address Count is display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			String[] List1 = Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText().split(" ");
			b1 = Integer.parseInt(List1[3]);
			System.out.println("Second Count ==>" + b1);
			getscreenshot("AddressBook25");
			Thread.sleep(5000);
		} else {
			System.out.println("Address Count is not display Proper as below.");
			System.out.println(Driver.findElement(By.xpath(
					"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div[5]/div/div[1]/div/div[2]/div[3]"))
					.getText());
			Thread.sleep(5000);
			getscreenshot("AddressBook26");
			Thread.sleep(5000);
		}

		c = a1 - a;
		d = b1 - b;
		System.out.println("First Count ==>" + c);
		System.out.println("Second Count ==>" + d);
		if (c > 0 || d > 0) {
			System.out.println("Address is properly ADDED in Address Book screen.");
			Thread.sleep(5000);
		} else {
			System.out.println("Address is properly ADDED in Address Book screen.");
			Thread.sleep(5000);
		}
	}

	@Test
	public void CreateASN() throws Exception {
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm");
		Date date = new Date();
		System.out.println("**********Create ASN**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("CreateASNTab")).click();
		Thread.sleep(20000);

		Select dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText("(select)");
		Thread.sleep(20000);

		Driver.findElement(By.id("hlkSaveASN")).click();
		Thread.sleep(5000);

		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("CreateASN1");
			Thread.sleep(5000);
			System.out.println("Create ASN FIRST validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Create ASN FIRST validation is not display proper.");
			getscreenshot("CreateASN2");
			Thread.sleep(5000);
		}

		dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameSPL);
		Thread.sleep(20000);

		Driver.findElement(By.id("hlkSaveASN")).click();
		Thread.sleep(5000);

		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("CreateASN3");
			Thread.sleep(5000);
			System.out.println("Create ASN SECOND validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Create ASN SECOND validation is not display proper.");
			getscreenshot("CreateASN4");
			Thread.sleep(5000);
		}

		dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameSPL);
		Thread.sleep(20000);

		Select dropdown2 = new Select(Driver.findElement(By.id("ddlfsl")));
		dropdown2.selectByVisibleText(FSLName1);
		Thread.sleep(20000);

		Driver.findElement(By.id("hlkSaveASN")).click();
		Thread.sleep(5000);

		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("CreateASN5");
			Thread.sleep(5000);
			System.out.println("Create ASN THIRD validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Create ASN THIRD validation is not display proper.");
			getscreenshot("CreateASN6");
			Thread.sleep(5000);
		}

		for (int i = 1; i <= 6; i++) {
			dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
			dropdown1.selectByVisibleText(CustomerNameSPL);
			Thread.sleep(20000);

			dropdown2 = new Select(Driver.findElement(By.id("ddlfsl")));
			dropdown2.selectByVisibleText(FSLName1);
			Thread.sleep(20000);

			Driver.findElement(By.id("txtestdate")).sendKeys(dateFormat.format(date));
			Thread.sleep(15000);

			Driver.findElement(By.id("AddNewNglId")).click();
			Thread.sleep(20000);

			Driver.findElement(By.id("txtNglFSL")).clear();
			Driver.findElement(By.id("txtNglFSL")).sendKeys(PartF1);
			Thread.sleep(5000);

			Driver.findElement(By.id("btnSearch")).click();
			Thread.sleep(20000);

			Driver.findElement(By.id("addparts_0")).click();
			Thread.sleep(15000);

			Driver.findElement(By.id("hlkSavePart")).click();
			Thread.sleep(20000);

			Select dropdown3 = new Select(Driver.findElement(By.id("ddlASNType")));
			dropdown3.selectByIndex(1);

			Driver.findElement(By.id("txtasnref")).sendKeys("ASNRefPDOSHI");
			Thread.sleep(5000);

			Select dropdown4 = new Select(Driver.findElement(By.id("ddlCarrier")));
			dropdown4.selectByIndex(i);
			Thread.sleep(5000);

			try {
				if (Driver.findElement(By.id("txtothercarriername")).isDisplayed() == true) {
					Driver.findElement(By.id("txtothercarriername")).sendKeys("PDOSHI");
					Thread.sleep(5000);
				}
			} catch (Exception e) {
				System.out.println("Carrier Name is ==> " + dropdown4.getFirstSelectedOption());
				Thread.sleep(5000);
			}

			Driver.findElement(By.id("txttrackingno")).sendKeys("ASNTrackPDOSHI");
			Thread.sleep(5000);

			Driver.findElement(By.id("txtnotes")).sendKeys("ASNNotePDOSHI");
			Thread.sleep(5000);

			Driver.findElement(By.id("imgExpand")).click();
			Thread.sleep(5000);

			try {
				Driver.findElement(By.id("AddNewNglChildPartid")).click();
				Thread.sleep(5000);
			} catch (Exception e) {
				JavascriptExecutor jse = (JavascriptExecutor) Driver;
				jse.executeScript("window.scrollBy(0,250)");
				Driver.findElement(By.id("AddNewNglChildPartid")).click();
				Thread.sleep(5000);
			}

			Driver.findElement(By.id("txtOrderQty")).clear();
			Driver.findElement(By.id("txtOrderQty")).sendKeys("1");
			Thread.sleep(5000);

			Driver.findElement(By.id("hlkSaveASN")).click();
			Thread.sleep(5000);

			if (Driver.findElement(By.id("success")).isDisplayed() == true) {
				getscreenshot("CreateASN7");
				Thread.sleep(5000);
				System.out.println("ASN is Create Successfully.");
				String Message1 = Driver.findElement(By.id("success")).getText();
				System.out.println(Message1);
				Thread.sleep(5000);
			} else {
				System.out.println("ASN is not Create Successfully.");
				getscreenshot("CreateASN8");
				Thread.sleep(5000);
			}
			try {
				Thread.sleep(5000);
				String ASNNo = Driver.findElement(By.id("txtasnno")).getText();
				System.out.println(ASNNo);
				Thread.sleep(5000);
				String WOASNNo = Driver.findElement(By.id("l_workOrderId")).getText();
				System.out.println(WOASNNo);
				Thread.sleep(5000);
			} catch (Exception e) {
				System.out.println("ASN is not Create.");
			}
			Driver.findElement(By.id("hlkCreateASN")).click();
			Thread.sleep(15000);
		}
	}

	@Test
	public void ASNLog() throws Exception {
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
		Calendar cal = Calendar.getInstance();
		System.out.println("**********ASN Log**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgasnlog")).click();
		Thread.sleep(20000);

		String Total = Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[3]/div/gridcontrol-controller/div/div/div[9]/div/div[1]"))
				.getText();
		Thread.sleep(5000);
		System.out.println(Total);

		Driver.findElement(By.id("cmbAsnType")).click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div[4]/div[1]/div/div/ul/li[1]/div/label/input"))
				.click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div[4]/div[1]/div/div/ul/li[1]/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnRunSearch")).click();
		Thread.sleep(15000);

		Total = Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[3]/div/gridcontrol-controller/div/div/div[9]/div/div[1]"))
				.getText();
		Thread.sleep(5000);
		System.out.println(Total);

		Driver.findElement(By.id("btnExport")).click();
		Thread.sleep(20000);

		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		getscreenshot("ASN LOG1");
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);

		// Estimate Date From and To
		Thread.sleep(5000);
		cal.add(Calendar.DATE, -30);
		Driver.findElement(By.id("txtFromEstArrival")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtFromEstArrival")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		Driver.findElement(By.id("txtToEstArrival")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtToEstArrival")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnRunSearch")).click();
		Thread.sleep(15000);

		Total = Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[3]/div/gridcontrol-controller/div/div/div[9]/div/div[1]"))
				.getText();
		Thread.sleep(5000);
		System.out.println(Total);
		System.out.println("ASN Log is working with 30 Days in Estimate Arrival Date.");

		Driver.findElement(By.id("btnExport")).click();
		Thread.sleep(20000);

		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		getscreenshot("ASN LOG2");
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);

		Thread.sleep(5000);
		cal.add(Calendar.DATE, -60);
		Driver.findElement(By.id("txtFromEstArrival")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtFromEstArrival")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		Driver.findElement(By.id("txtToEstArrival")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtToEstArrival")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnRunSearch")).click();
		Thread.sleep(15000);

		Total = Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[3]/div/gridcontrol-controller/div/div/div[9]/div/div[1]"))
				.getText();
		Thread.sleep(5000);
		System.out.println(Total);
		System.out.println("ASN Log is working with 60 Days in Estimate Arrival Date.");

		Driver.findElement(By.id("btnExport")).click();
		Thread.sleep(20000);

		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		getscreenshot("ASN LOG3");
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);

		// Clear Estimate From and To
		Driver.findElement(By.id("txtFromEstArrival")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtToEstArrival")).clear();
		Thread.sleep(5000);

		// ASN Date From and To
		Thread.sleep(5000);
		cal.add(Calendar.DATE, -30);
		Driver.findElement(By.id("txtAsnFromDate")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtAsnFromDate")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		Driver.findElement(By.id("txtAsnToDate")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtAsnToDate")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnRunSearch")).click();
		Thread.sleep(15000);

		Total = Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[3]/div/gridcontrol-controller/div/div/div[9]/div/div[1]"))
				.getText();
		Thread.sleep(5000);
		System.out.println(Total);
		System.out.println("ASN Log is working with 30 Days in ASN Date.");

		Driver.findElement(By.id("btnExport")).click();
		Thread.sleep(20000);

		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		getscreenshot("ASN LOG4");
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);

		Thread.sleep(5000);
		cal.add(Calendar.DATE, -60);
		Driver.findElement(By.id("txtAsnFromDate")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtAsnFromDate")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		Driver.findElement(By.id("txtAsnToDate")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtAsnToDate")).sendKeys(dateFormat.format(cal.getTime()));
		Thread.sleep(5000);
		Driver.findElement(By.id("txtWorkOrder")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnRunSearch")).click();
		Thread.sleep(15000);

		Total = Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div/div[2]/div/div[3]/div/gridcontrol-controller/div/div/div[9]/div/div[1]"))
				.getText();
		Thread.sleep(5000);
		System.out.println(Total);
		System.out.println("ASN Log is working with 60 Days in ASN Date.");

		Driver.findElement(By.id("btnExport")).click();
		Thread.sleep(20000);

		((JavascriptExecutor) Driver).executeScript("window.scrollTo(0,document.body.scrollHeight)");
		Thread.sleep(5000);
		getscreenshot("ASN LOG5");
		Thread.sleep(5000);
		((JavascriptExecutor) Driver).executeScript("window.scrollTo(document.body.scrollHeight,0)");
		Thread.sleep(5000);
	}

	@Test
	public void ReorderLog() throws Exception {
		Random rand = new Random();
		System.out.println("**********Reorder Log**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgReorderlog")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(5000);

		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("ReorderLOG1");
			Thread.sleep(5000);
			System.out.println("Reorder LOG FIRST validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Reorder LOG FIRST validation is not display proper.");
			getscreenshot("ReorderLOG2");
			Thread.sleep(5000);
		}

		Select dropdown1 = new Select(Driver.findElement(By.id("ddlClient")));
		dropdown1.selectByVisibleText(CustomerNameSPL);
		Thread.sleep(20000);

		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(5000);

		if (Driver.findElement(By.id("idValidation")).isDisplayed() == true) {
			getscreenshot("ReorderLOG3");
			Thread.sleep(5000);
			System.out.println("Reorder LOG Second validation is display proper.");
			String Message1 = Driver.findElement(By.id("idValidation")).getText();
			System.out.println(Message1);
			Thread.sleep(5000);
		} else {
			System.out.println("Reorder LOG Second validation is not display proper.");
			getscreenshot("ReorderLOG4");
			Thread.sleep(5000);
		}

		Select dropdown2 = new Select(Driver.findElement(By.id("ddlfsl")));
		dropdown2.selectByVisibleText(FSLName1);
		Thread.sleep(20000);

		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);

		getscreenshot("ReorderLOG5");
		Thread.sleep(5000);

		String num1 = Integer.toString(rand.nextInt(20));
		String num2 = Integer.toString(rand.nextInt(99));

		for (int i = 1; i <= 6; i++) {
			Select dropdown3 = new Select(Driver.findElement(By.id("ddlClassCode")));
			if (i == 1) {
				dropdown3.selectByVisibleText("A");
				Thread.sleep(15000);

				Driver.findElement(By.id("btnSearch")).click();
				Thread.sleep(5000);

				getscreenshot("ReorderLOG6A" + i);
				Thread.sleep(5000);

				for (int j = 0; j <= 10; j++) {
					String fieldname = "txtReorderPoint_" + j;
					Driver.findElement(By.id(fieldname)).clear();
					Driver.findElement(By.id(fieldname)).sendKeys(num1);
					Thread.sleep(5000);

					String fieldname1 = "txtMinOrderQuantity_" + j;
					Driver.findElement(By.id(fieldname1)).clear();
					Driver.findElement(By.id(fieldname1)).sendKeys(num2);
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("hlkSaveReorder")).click();
				Thread.sleep(20000);

				if (Driver.findElement(By.id("success")).isDisplayed() == true) {
					getscreenshot("ReorderLOG7");
					Thread.sleep(5000);
					System.out.println("Reorder Data are save Successfully.");
					String Message1 = Driver.findElement(By.id("success")).getText();
					System.out.println(Message1);
					Thread.sleep(5000);
				} else {
					System.out.println("Reorder Data are not save Successfully.");
					getscreenshot("ReorderLOG8");
					Thread.sleep(5000);
				}

				getscreenshot("ReorderLOG9A" + i);
				Thread.sleep(5000);
			} else if (i == 2) {
				dropdown3.selectByVisibleText("B");
				Thread.sleep(15000);

				Driver.findElement(By.id("btnSearch")).click();
				Thread.sleep(5000);

				getscreenshot("ReorderLOG10B" + i);
				Thread.sleep(5000);

				for (int j = 0; j <= 10; j++) {
					String fieldname = "txtReorderPoint_" + j;
					Driver.findElement(By.id(fieldname)).clear();
					Driver.findElement(By.id(fieldname)).sendKeys(num1);
					Thread.sleep(5000);

					String fieldname1 = "txtMinOrderQuantity_" + j;
					Driver.findElement(By.id(fieldname1)).clear();
					Driver.findElement(By.id(fieldname1)).sendKeys(num2);
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("hlkSaveReorder")).click();
				Thread.sleep(20000);

				if (Driver.findElement(By.id("success")).isDisplayed() == true) {
					getscreenshot("ReorderLOG11");
					Thread.sleep(5000);
					System.out.println("Reorder Data are save Successfully.");
					String Message1 = Driver.findElement(By.id("success")).getText();
					System.out.println(Message1);
					Thread.sleep(5000);
				} else {
					System.out.println("Reorder Data are not save Successfully.");
					getscreenshot("ReorderLOG12");
					Thread.sleep(5000);
				}

				getscreenshot("ReorderLOG13B" + i);
				Thread.sleep(5000);
			} else if (i == 3) {
				dropdown3.selectByVisibleText("C");
				Thread.sleep(15000);

				Driver.findElement(By.id("btnSearch")).click();
				Thread.sleep(5000);

				getscreenshot("ReorderLOG14C" + i);
				Thread.sleep(5000);

				for (int j = 0; j <= 10; j++) {
					String fieldname = "txtReorderPoint_" + j;
					Driver.findElement(By.id(fieldname)).clear();
					Driver.findElement(By.id(fieldname)).sendKeys(num1);
					Thread.sleep(5000);

					String fieldname1 = "txtMinOrderQuantity_" + j;
					Driver.findElement(By.id(fieldname1)).clear();
					Driver.findElement(By.id(fieldname1)).sendKeys(num2);
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("hlkSaveReorder")).click();
				Thread.sleep(20000);

				if (Driver.findElement(By.id("success")).isDisplayed() == true) {
					getscreenshot("ReorderLOG15");
					Thread.sleep(5000);
					System.out.println("Reorder Data are save Successfully.");
					String Message1 = Driver.findElement(By.id("success")).getText();
					System.out.println(Message1);
					Thread.sleep(5000);
				} else {
					System.out.println("Reorder Data are not save Successfully.");
					getscreenshot("ReorderLOG16");
					Thread.sleep(5000);
				}

				getscreenshot("ReorderLOG17B" + i);
				Thread.sleep(5000);
			} else if (i == 4) {
				dropdown3.selectByVisibleText("D");
				Thread.sleep(15000);

				Driver.findElement(By.id("btnSearch")).click();
				Thread.sleep(5000);

				getscreenshot("ReorderLOG18D" + i);
				Thread.sleep(5000);

				for (int j = 0; j <= 10; j++) {
					String fieldname = "txtReorderPoint_" + j;
					Driver.findElement(By.id(fieldname)).clear();
					Driver.findElement(By.id(fieldname)).sendKeys(num1);
					Thread.sleep(5000);

					String fieldname1 = "txtMinOrderQuantity_" + j;
					Driver.findElement(By.id(fieldname1)).clear();
					Driver.findElement(By.id(fieldname1)).sendKeys(num2);
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("hlkSaveReorder")).click();
				Thread.sleep(20000);

				if (Driver.findElement(By.id("success")).isDisplayed() == true) {
					getscreenshot("ReorderLOG19");
					Thread.sleep(5000);
					System.out.println("Reorder Data are save Successfully.");
					String Message1 = Driver.findElement(By.id("success")).getText();
					System.out.println(Message1);
					Thread.sleep(5000);
				} else {
					System.out.println("Reorder Data are not save Successfully.");
					getscreenshot("ReorderLOG20");
					Thread.sleep(5000);
				}

				getscreenshot("ReorderLOG21A" + i);
				Thread.sleep(5000);
			} else if (i == 5) {
				dropdown3.selectByVisibleText("E");
				Thread.sleep(15000);

				Driver.findElement(By.id("btnSearch")).click();
				Thread.sleep(5000);

				getscreenshot("ReorderLOG22E" + i);
				Thread.sleep(5000);

				for (int j = 0; j <= 10; j++) {
					String fieldname = "txtReorderPoint_" + j;
					Driver.findElement(By.id(fieldname)).clear();
					Driver.findElement(By.id(fieldname)).sendKeys(num1);
					Thread.sleep(5000);

					String fieldname1 = "txtMinOrderQuantity_" + j;
					Driver.findElement(By.id(fieldname1)).clear();
					Driver.findElement(By.id(fieldname1)).sendKeys(num2);
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("hlkSaveReorder")).click();
				Thread.sleep(20000);

				if (Driver.findElement(By.id("success")).isDisplayed() == true) {
					getscreenshot("ReorderLOG23");
					Thread.sleep(5000);
					System.out.println("Reorder Data are save Successfully.");
					String Message1 = Driver.findElement(By.id("success")).getText();
					System.out.println(Message1);
					Thread.sleep(5000);
				} else {
					System.out.println("Reorder Data are not save Successfully.");
					getscreenshot("ReorderLOG24");
					Thread.sleep(5000);
				}

				getscreenshot("ReorderLOG25E" + i);
				Thread.sleep(5000);
			} else if (i == 6) {
				dropdown3.selectByVisibleText("F");
				Thread.sleep(15000);

				Driver.findElement(By.id("btnSearch")).click();
				Thread.sleep(5000);

				getscreenshot("ReorderLOG26F" + i);
				Thread.sleep(5000);

				for (int j = 0; j <= 10; j++) {
					String fieldname = "txtReorderPoint_" + j;
					Driver.findElement(By.id(fieldname)).clear();
					Driver.findElement(By.id(fieldname)).sendKeys(num1);
					Thread.sleep(5000);

					String fieldname1 = "txtMinOrderQuantity_" + j;
					Driver.findElement(By.id(fieldname1)).clear();
					Driver.findElement(By.id(fieldname1)).sendKeys(num2);
					Thread.sleep(5000);
				}

				Driver.findElement(By.id("hlkSaveReorder")).click();
				Thread.sleep(20000);

				if (Driver.findElement(By.id("success")).isDisplayed() == true) {
					getscreenshot("ReorderLOG27");
					Thread.sleep(5000);
					System.out.println("Reorder Data are save Successfully.");
					String Message1 = Driver.findElement(By.id("success")).getText();
					System.out.println(Message1);
					Thread.sleep(5000);
				} else {
					System.out.println("Reorder Data are not save Successfully.");
					getscreenshot("ReorderLOG28");
					Thread.sleep(5000);
				}

				getscreenshot("ReorderLOG29F" + i);
				Thread.sleep(5000);
			} else {
				dropdown3.selectByVisibleText("There is no value in Combo.");
				Thread.sleep(15000);

				getscreenshot("ReorderLOG30");
				Thread.sleep(5000);
			}
		}
	}

	@Test
	public void FlightView() throws Exception {
		Robot robot = new Robot();
		System.out.println("**********Flight View**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("imgFlightview")).click();
		Thread.sleep(20000);

		getscreenshot("FlightView1");
		Thread.sleep(5000);

		Driver.findElement(By.id("btnSearchCharges")).click();
		Thread.sleep(15000);

		getscreenshot("FlightView2");
		Thread.sleep(5000);

		Driver.findElement(By.xpath("/html/body/div[3]/section/div[2]/div[1]/div[1]/div/div[4]/div/div")).click();
		Thread.sleep(5000);

		Driver.findElement(By
				.xpath("/html/body/div[3]/section/div[2]/div[1]/div[1]/div/div[4]/div/div/ul/li[4]/a/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(
				By.xpath("/html/body/div[3]/section/div[2]/div[1]/div[1]/div/div[4]/div/div/ul/li[1]/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnSearchCharges")).click();
		Thread.sleep(15000);

		getscreenshot("FlightView3");
		Thread.sleep(5000);

		Driver.findElement(By.xpath("/html/body/div[3]/section/div[2]/div[1]/div[1]/div/div[4]/div/div")).click();
		Thread.sleep(5000);

		Driver.findElement(
				By.xpath("/html/body/div[3]/section/div[2]/div[1]/div[1]/div/div[4]/div/div/ul/li[1]/div/label/input"))
				.click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnSearchCharges")).click();
		Thread.sleep(15000);

		getscreenshot("FlightView4");
		Thread.sleep(5000);

		Driver.findElement(By.id("txtOriginAirport")).clear();
		Thread.sleep(5000);

		Driver.findElement(By.id("txtOriginAirport")).sendKeys("DCA");
		Thread.sleep(15000);
		robot.keyPress(KeyEvent.VK_DOWN);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtDestAirport")).clear();
		Thread.sleep(5000);

		Driver.findElement(By.id("txtDestAirport")).sendKeys("ATL");
		Thread.sleep(15000);
		robot.keyPress(KeyEvent.VK_DOWN);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);

		Driver.findElement(By.id("btnSearchCharges")).click();
		Thread.sleep(15000);

		getscreenshot("FlightView5");
		Thread.sleep(5000);

		Driver.findElement(By.id("txtOriginAirport")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtDestAirport")).clear();
		Thread.sleep(5000);

		Driver.findElement(By.id("txtAirline")).clear();
		Thread.sleep(5000);
		Driver.findElement(By.id("txtAirline")).sendKeys("DELTA AIRLINES");
		Thread.sleep(15000);
		robot.keyPress(KeyEvent.VK_DOWN);
		robot.keyPress(KeyEvent.VK_ENTER);
		Thread.sleep(5000);

		Driver.findElement(By.id("btnSearchCharges")).click();
		Thread.sleep(15000);

		getscreenshot("FlightView6");
		Thread.sleep(5000);
	}

	@Test
	public void Temp1() throws Exception {
		Calendar cal = Calendar.getInstance();
		DateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy hh:mm");
		System.out.println("Today's date is " + dateFormat.format(cal.getTime()));

		cal.add(Calendar.DATE, -1);
		System.out.println("Yesterday's date was " + dateFormat.format(cal.getTime()));

		cal.add(Calendar.DATE, -30);
		System.out.println("Yesterday's date was " + dateFormat.format(cal.getTime()));
	}

	@Test
	public void ConsigneeAddrImport() throws Exception {
		Robot robot = new Robot();
		System.out.println("**********Consignee Addr Import**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idConsigneeAddrImport")).click();
		Thread.sleep(20000);
		Driver.findElement(By.id("hlkImportConsigneeAddr")).click();
		Thread.sleep(20000);
		Driver.findElement(By.id("file")).sendKeys(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Consignee Address.xls");
		Thread.sleep(5000);
		Driver.findElement(By.id("btnUpload")).click();
		Thread.sleep(20000);

		String Message1 = Driver.findElement(By.id("successid")).getText();

		if (Message1.equals("Import Process Completed !")) {
			SheetMessage = "*****Import Process is Completed !*****";
			System.out.println(SheetMessage);
			Thread.sleep(5000);
		} else {
			Message1 = Driver.findElement(By.id("errorid")).getText();
			System.out.println(Message1);
			SheetMessage = "*****Import Process is not Completed !*****";
			System.out.println(SheetMessage);
			Thread.sleep(5000);
		}
		Thread.sleep(5000);
		robot.keyPress(KeyEvent.VK_ESCAPE);
		Thread.sleep(5000);
		try {
			Driver.findElement(By.id("hlkViewSampleFile")).click();
			Thread.sleep(5000);
		} catch (Exception e) {
			JavascriptExecutor jse = (JavascriptExecutor) Driver;
			jse.executeScript("window.scrollBy(0,250)");
			Driver.findElement(By.id("hlkViewSampleFile")).click();
			Thread.sleep(5000);
		}
		try {
			Driver.findElement(By.id("hlkExportConsigneeAddress")).click();
			Thread.sleep(5000);
		} catch (Exception e) {
			JavascriptExecutor jse = (JavascriptExecutor) Driver;
			jse.executeScript("window.scrollBy(0,250)");
			Driver.findElement(By.id("hlkExportConsigneeAddress")).click();
			Thread.sleep(5000);
		}
	}

	@Test
	public void ViewInvoice() throws Exception {
		System.out.println("**********View Invoice**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idViewInvoice")).click();
		Thread.sleep(20000);
		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);
		Driver.findElement(By.id("btnReset")).click();
		Thread.sleep(20000);
		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(5000);

		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipViewInvoiceDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipViewInvoiceSTG.xlsx");

		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("Sheet1");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();

		System.out.println("Row Count ====> " + rcount);

		String Message1 = Driver.findElement(By.id("errorid")).getText();

		if (Message1.equals("Please provide search criteria.")) {
			SheetMessage = "*****Need to Add From and To Date for Search Invoice.*****";
			System.out.println(SheetMessage);
			Thread.sleep(5000);
		} else {
			SheetMessage = "*****Validation Message is not display Proper.*****";
			System.out.println(SheetMessage);
			System.out.println("It is display This Message ==> " + Message1);
			Thread.sleep(5000);
		}
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -60);
		String ValiFrom = getDate(cal);
		Thread.sleep(5000);
		System.out.println("Valid From Date :- " + ValiFrom);
		String ValiTo = CuDate();
		System.out.println("Valid To Date :- " + ValiTo);

		Driver.findElement(By.id("txtValidFrom")).sendKeys(ValiFrom);
		Driver.findElement(By.id("txtValidTo")).sendKeys(ValiTo);

		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("btnSend")).click();
		Thread.sleep(5000);

		String Message2 = Driver.findElement(By.id("errorid")).getText();

		if (Message2.equals("Please enter valid email.")) {
			SheetMessage = "*****Need to Add E-Mail Id for Send Invoice.*****";
			System.out.println(SheetMessage);
			Thread.sleep(5000);
		} else {
			SheetMessage = "*****Validation Message is not display Proper.*****";
			System.out.println(SheetMessage);
			System.out.println("It is display This Message ==> " + Message2);
			Thread.sleep(5000);
		}
		Driver.findElement(By.id("txtAddress")).sendKeys("pdoshi@samyak.com");
		Driver.findElement(By.id("btnSend")).click();
		Thread.sleep(5000);

		String Message3 = Driver.findElement(By.id("errorid")).getText();

		if (Message3.equals("Please select atleast one invoice for send Invoice.")) {
			SheetMessage = "*****Need to select atleast one Invoice for Send Invoice.*****";
			System.out.println(SheetMessage);
			Thread.sleep(5000);
		} else {
			SheetMessage = "*****Validation Message is not display Proper.*****";
			System.out.println(SheetMessage);
			System.out.println("It is display This Message ==> " + Message3);
			Thread.sleep(5000);
		}

		Driver.findElement(By.id("btnReset")).click();
		Thread.sleep(20000);

		Select dropdown = new Select(Driver.findElement(By.id("drpClient")));
		dropdown.selectByVisibleText(formatter.formatCellValue(sh0.getRow(1).getCell(2)));

		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);

		String Message4 = Driver.findElement(By.id("errorid")).getText();

		if (Message4.equals("Please provide Invoice From and To Date.")) {
			SheetMessage = "*****Need to Add From and To Date for Search Invoice.*****";
			System.out.println(SheetMessage);
			Thread.sleep(5000);
		} else {
			SheetMessage = "*****Validation Message is not display Proper.*****";
			System.out.println(SheetMessage);
			System.out.println("It is display This Message ==> " + Message4);
			Thread.sleep(5000);
		}

		Driver.findElement(By.id("btnReset")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("txtinvoicenumber")).sendKeys(formatter.formatCellValue(sh0.getRow(2).getCell(0)));
		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);

		String Values1 = Driver.findElement(By.id("SearchInvoice")).getText();
		System.out.println("ALL DATA ====> " + Values1);
		String[] List1 = Values1.split(" ");
		System.out.println("After Split DATA ====> " + List1[43]);
		String Values2 = List1[41];
		System.out.println("After Split DATA ====> " + List1[41]);

		if (Values2.equals("Paid")) {
			System.out.println("Entered Invoice is Paid.");
		} else {
			System.out.println("Entered Invoice is Unpaid.");
		}
		Thread.sleep(15000);

		Driver.findElement(By.id("txtinvoicenumber")).clear();
		Driver.findElement(By.id("txtinvoicenumber")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(0)));
		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);

		String Values3 = Driver.findElement(By.id("SearchInvoice")).getText();
		System.out.println("ALL DATA ====> " + Values3);
		String[] List2 = Values3.split(" ");
		System.out.println("After Split DATA ====> " + List2[43]);
		String Values4 = List2[41];
		System.out.println("After Split DATA ====> " + List2[41]);

		if (Values4.equals("Paid")) {
			System.out.println("Entered Invoice is Paid.");
		} else {
			System.out.println("Entered Invoice is Unpaid.");
		}
		Thread.sleep(15000);

		String SearchInvoice = "hlkDetails" + formatter.formatCellValue(sh0.getRow(1).getCell(0));
		System.out.println(SearchInvoice);

		Driver.findElement(By.id(SearchInvoice)).click();
		Thread.sleep(20000);

		String SearchView = "hrefCView" + formatter.formatCellValue(sh0.getRow(1).getCell(3));
		System.out.println(SearchView);

		Driver.findElement(By.id(SearchView)).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("backlist")).click();
		Thread.sleep(20000);

		String SearchInvPrint = "hlkViewPrint" + formatter.formatCellValue(sh0.getRow(1).getCell(0));

		Driver.findElement(By.id(SearchInvPrint)).click();
		Thread.sleep(20000);

		String strParentWindowHandle = Driver.getWindowHandle();

		String ViewPrintURL = Driver.getCurrentUrl();
		System.out.println("After Click on View Print This URL Display ====> " + ViewPrintURL);
		Thread.sleep(5000);

		for (String winHandle : Driver.getWindowHandles()) {
			Driver.switchTo().window(winHandle);
		}

		Driver.close();

		// Switch back to original browser (first window)
		Driver.switchTo().window(strParentWindowHandle);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtAddress")).sendKeys("pdoshi@samyak.com");

		String CheckInvBox = "chkInvoice" + formatter.formatCellValue(sh0.getRow(1).getCell(0));

		Driver.findElement(By.id(CheckInvBox)).click();
		Thread.sleep(5000);

		WebElement CheckBox1 = Driver.findElement(By.id("chkExcludePaidInv"));
		if (CheckBox1.isSelected()) {
			Driver.findElement(By.id("chkExcludePaidInv")).click();
			Thread.sleep(5000);
		}

		WebElement CheckBox2 = Driver.findElement(By.id("chkExcludeNGLLogo"));
		if (CheckBox2.isSelected()) {
			Driver.findElement(By.id("chkExcludePaidInv")).click();
			Thread.sleep(5000);
		}

		Driver.findElement(By.id("btnSend")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("btnReset")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("txtinvoicenumber")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(0)));
		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("txtAddress")).sendKeys("pdoshi@samyak.com");

		Driver.findElement(By.id(CheckInvBox)).click();
		Thread.sleep(5000);

		WebElement CheckBox3 = Driver.findElement(By.id("chkExcludePaidInv"));
		if (!CheckBox3.isSelected()) {
			Driver.findElement(By.id("chkExcludePaidInv")).click();
			Thread.sleep(5000);
		}

		WebElement CheckBox4 = Driver.findElement(By.id("chkExcludeNGLLogo"));
		if (!CheckBox4.isSelected()) {
			Driver.findElement(By.id("chkExcludePaidInv")).click();
			Thread.sleep(5000);
		}

		Driver.findElement(By.id("btnSend")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("btnReset")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("txtinvoicenumber")).sendKeys(formatter.formatCellValue(sh0.getRow(1).getCell(0)));
		Driver.findElement(By.id("btnSearch")).click();
		Thread.sleep(20000);

		String Values5 = Driver.findElement(By.id("SearchInvoice")).getText();
		System.out.println("ALL DATA ====> " + Values5);
		String[] List3 = Values5.split(" ");
		String Values7 = CuDate();
		String Values6 = List3[42];
		System.out.println("After Split DATA ====> " + List3[42]);

		String SendButton1 = "btnSendInvoice" + formatter.formatCellValue(sh0.getRow(1).getCell(0));

		Driver.findElement(By.id(SendButton1)).click();
		Thread.sleep(5000);

		Values5 = Driver.findElement(By.id("SearchInvoice")).getText();
		System.out.println("ALL DATA ====> " + Values5);
		List3 = Values5.split(" ");
		Values7 = CuDate().trim();
		Values6 = List3[42].trim();
		System.out.println("After Split DATA ====> " + List3[42]);

		if (Values6.equals(Values7)) {
			System.out.println("Sent Date is display Proper.");
			System.out.println("Sent Date ==> " + Values6);
			System.out.println("Todays Date ==> " + Values7);
		} else {
			System.out.println("Sent Date is not display Proper.");
			System.out.println("Sent Date ==> " + Values6);
			System.out.println("Todays Date ==> " + Values7);
		}
		Thread.sleep(15000);
	}

	@Test
	public void LookUpShipment() throws Exception {
		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipLookUpShipmentDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipLookUpShipmentSTG.xlsx");

		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("LookUpShipment");
		int rcount = sh0.getLastRowNum();
		System.out.println("**********Look Up Shipment**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idLookUpShipment")).click();
		Thread.sleep(20000);

		Driver.findElement(By.id("btnTrackOrder")).click();

		String Message1 = Driver.findElement(By.id("errorid")).getText();

		if (Message1.equals("Please enter value in any one of the field.")) {
			System.out.println("*****VALIDATION IS DISPLAY PROPER ON SCREEN.*****");
			System.out.println(Message1);
		} else {
			System.out.println("*****VALIDATION IS NOT DISPLAY PROPER ON SCREEN.*****");
			System.out.println(Message1);
		}
		Thread.sleep(5000);

		getscreenshot("LookUpShipment1");

		System.out.println("Row Count ====> " + rcount);

		for (int i = 1; i <= rcount; i++) {
			System.out.println(
					"****************************************************************************************************");
			DataFormatter formatter = new DataFormatter();
			try {
				Driver.findElement(By.id("txtPickupId")).clear();
				Driver.findElement(By.id("txtPickupId")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(0)));
				Driver.findElement(By.id("btnTrackOrder")).click();
				Thread.sleep(5000);

				String Message2 = Driver.findElement(By.id("errorid")).getText();

				if (Message2.equals("No Record Found.")) {
					Driver.findElement(By.id("txtPickupId")).clear();
					System.out.println("*****There is no Job with "
							+ formatter.formatCellValue(sh0.getRow(i).getCell(0)) + "*****");
					System.out.println(Message2);
					Thread.sleep(5000);
					getscreenshot("LookUpShipment2");
				} else {
					Thread.sleep(15000);
					Driver.findElement(By.id("hlkBackToScreen")).click();
					System.out.println("Entered Job's Shipment Detail Page is Display Proper.");
					Thread.sleep(5000);
				}
			} catch (Exception e) {
				System.out.println(
						"Please check This Pickup Manualy ==> " + formatter.formatCellValue(sh0.getRow(i).getCell(0)));
			}

			try {
				Driver.findElement(By.id("txtPickupId")).clear();
				Driver.findElement(By.id("txtBOL")).clear();
				Driver.findElement(By.id("txtBOL")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(1)));
				Driver.findElement(By.id("btnTrackOrder")).click();
				Thread.sleep(5000);

				String Message3 = Driver.findElement(By.id("errorid")).getText();

				if (Message3.equals("No Record Found.")) {
					Driver.findElement(By.id("txtBOL")).clear();
					System.out.println("*****There is no Job with "
							+ formatter.formatCellValue(sh0.getRow(i).getCell(1)) + "*****");
					System.out.println(Message3);
					Thread.sleep(5000);
					getscreenshot("LookUpShipment3");
				} else {
					Thread.sleep(20000);
					Driver.findElement(By.id("hlkBackToScreen")).click();
					System.out.println("Entered Job's Shipment Detail Page is Display Proper.");
					Thread.sleep(5000);
				}
			} catch (Exception e) {
				System.out.println(
						"Please check This BOL Manualy ==> " + formatter.formatCellValue(sh0.getRow(i).getCell(1)));
			}

			try {
				Driver.findElement(By.id("txtBOL")).clear();
				Driver.findElement(By.id("txtReference")).clear();
				Driver.findElement(By.id("txtReference")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(2)));
				Driver.findElement(By.id("btnTrackOrder")).click();
				Thread.sleep(5000);

				try {
					Thread.sleep(20000);
					WebElement ToGetRows = Driver.findElement(By.id("invoicedetailtable"));
					List<WebElement> TotalRowsList = ToGetRows.findElements(By.tagName("tr"));
					int RowCount1 = TotalRowsList.size();
					int RowCount2 = RowCount1 - 1;
					System.out.println("Total number of Rows in the table are : " + RowCount2);
					Thread.sleep(5000);

					WebElement ToGetColumns = Driver.findElement(By.id("invoicedetailtable"));
					List<WebElement> TotalColsList = ToGetColumns.findElements(By.tagName("td"));
					System.out.println("Total Number of Columns in the table are: " + TotalColsList.size());
					Thread.sleep(5000);

					Driver.findElement(By.id("hlkBackToScreen")).click();
					Thread.sleep(5000);
					System.out.println("Reference have Multiple Records.");
				} catch (Exception e) {
					String Message4 = Driver.findElement(By.id("errorid")).getText();

					if (Message4.equals("No Record Found.")) {
						Driver.findElement(By.id("txtReference")).clear();
						System.out.println("*****There is no Job with "
								+ formatter.formatCellValue(sh0.getRow(i).getCell(2)) + "*****");
						System.out.println(Message4);
						Thread.sleep(5000);
						getscreenshot("LookUpShipment4");
					} else {
						Thread.sleep(20000);
						Driver.findElement(By.id("hlkBackToScreen")).click();
						System.out.println("Entered Job's Shipment Detail Page is Display Proper.");
						Thread.sleep(5000);
					}
				}
			} catch (Exception e) {
				System.out.println("Please check This Reference Manualy ==> "
						+ formatter.formatCellValue(sh0.getRow(i).getCell(2)));
			}

			try {
				Driver.findElement(By.id("txtReference")).clear();
				Driver.findElement(By.id("txtJobId")).clear();
				Driver.findElement(By.id("txtJobId")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(3)));
				Driver.findElement(By.id("btnTrackOrder")).click();
				Thread.sleep(5000);

				String Message5 = Driver.findElement(By.id("errorid")).getText();

				if (Message5.equals("No Record Found.")) {
					Driver.findElement(By.id("txtJobId")).clear();
					System.out.println("*****There is no Job with "
							+ formatter.formatCellValue(sh0.getRow(i).getCell(3)) + "*****");
					System.out.println(Message5);
					Thread.sleep(5000);
					getscreenshot("LookUpShipment5");
				} else {
					Thread.sleep(20000);
					Driver.findElement(By.id("hlkBackToScreen")).click();
					System.out.println("Entered Job's Shipment Detail Page is Display Proper.");
					Thread.sleep(5000);
				}
			} catch (Exception e) {
				System.out.println(
						"Please check This Job Manualy ==> " + formatter.formatCellValue(sh0.getRow(i).getCell(3)));
			}

			Driver.findElement(By.id("txtJobId")).clear();
			Driver.findElement(By.id("txtOrderBy")).clear();
			Driver.findElement(By.id("txtOrderBy")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(4)));
			Driver.findElement(By.id("btnTrackOrder")).click();
			Thread.sleep(5000);

			try {
				Thread.sleep(20000);
				WebElement ToGetRows = Driver.findElement(By.id("invoicedetailtable"));
				List<WebElement> TotalRowsList = ToGetRows.findElements(By.tagName("tr"));
				int RowCount1 = TotalRowsList.size();
				int RowCount2 = RowCount1 - 1;
				System.out.println("Total number of Rows in the table are : " + RowCount2);
				Thread.sleep(5000);

				WebElement ToGetColumns = Driver.findElement(By.id("invoicedetailtable"));
				List<WebElement> TotalColsList = ToGetColumns.findElements(By.tagName("td"));
				System.out.println("Total Number of Columns in the table are: " + TotalColsList.size());
				Thread.sleep(5000);

				Driver.findElement(By.id("hlkBackToScreen")).click();
				Thread.sleep(5000);
				System.out.println("Caller Name have Multiple Records.");
			} catch (Exception e) {
				String Message6 = Driver.findElement(By.id("errorid")).getText();

				if (Message6.equals("No Record Found.")) {
					Driver.findElement(By.id("txtOrderBy")).clear();
					System.out.println("*****There is no Job with "
							+ formatter.formatCellValue(sh0.getRow(i).getCell(4)) + "*****");
					System.out.println(Message6);
					Thread.sleep(5000);
					getscreenshot("LookUpShipment6");
				} else {
					try {
						Thread.sleep(20000);
						Driver.findElement(By.id("hlkBackToScreen")).click();
						System.out.println("Entered Job's Shipment Detail Page is Display Proper.");
						Thread.sleep(5000);
						System.out.println("Caller Name have only single Record.");
						System.out.println("Entered Job's Shipment Detail Page is Display Proper.");
					} catch (Exception f) {
						System.out.println("Please check This Caller Name Manualy ==> "
								+ formatter.formatCellValue(sh0.getRow(i).getCell(4)));
					}
				}
				Driver.findElement(By.id("txtOrderBy")).clear();
			}
			System.out.println(
					"****************************************************************************************************");
		}
	}

	@Test
	public void ManageBatchOrders() throws Exception {
		Robot robot = new Robot();
		System.out.println("**********Manage Batch Orders**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idViewbatch")).click();
		Thread.sleep(20000);

		getscreenshot("ManageBatchOrders1");

		System.out.println(Driver
				.findElement(
						By.xpath("/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div[3]/label"))
				.getText());
		Thread.sleep(5000);

		Driver.findElement(By.id("btnCreateBatch")).click();
		Thread.sleep(15000);

		getscreenshot("ManageBatchOrders2");

		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);

		getscreenshot("ManageBatchOrders3");

		String MBOV1 = Driver.findElement(By.id("idValidation")).getText();
		System.out.println("Manage Batch Orders Validation : " + MBOV1);
		Thread.sleep(5000);

		Driver.findElement(By.id("hlkBackToScreen")).click();
		Thread.sleep(5000);

		getscreenshot("ManageBatchOrders4");

		Driver.findElement(By.id("btnCreateBatch")).click();
		Thread.sleep(15000);

		Select dropdown = new Select(Driver.findElement(By.id("drpcustomer")));
		dropdown.selectByVisibleText(CustomerNameNSPL);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtorderby")).sendKeys("Automation");
		Thread.sleep(2000);

		Driver.findElement(By.id("txtorderphone")).sendKeys("111-213-1415");
		Thread.sleep(2000);

		Driver.findElement(By.id("txtcompany")).sendKeys("Automation Company");
		Thread.sleep(2000);

		Driver.findElement(By.id("txtPUZipCode")).sendKeys("21225");
		robot.keyPress(KeyEvent.VK_TAB);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtaddressline")).sendKeys("101, Automation App.");
		Thread.sleep(2000);

		Driver.findElement(By.xpath(
				"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div[2]/div[2]/div[2]/div[5]/div/div[2]/span"))
				.click();
		Thread.sleep(2000);

		Driver.findElement(By.id("txtrdydate")).sendKeys(CuDateTime());
		Thread.sleep(2000);

		Driver.findElement(By.id("hlkAdd")).click();
		Thread.sleep(15000);

		Driver.findElement(By.id("txtdlcompany")).sendKeys("Automation Delivery Company");
		Thread.sleep(2000);

		Driver.findElement(By.id("txtDLZipCode")).sendKeys("21227");
		robot.keyPress(KeyEvent.VK_TAB);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtdladdressline")).sendKeys("102, Automation App.");
		Thread.sleep(2000);

		Select dropdown2 = new Select(Driver.findElement(By.id("cmbdefualtservice")));
		dropdown2.selectByIndex(1);
		Thread.sleep(5000);

		Driver.findElement(By.id("hrefAddNew")).click();
		Thread.sleep(2000);
		Driver.findElement(By.id("hrefAddNew")).click();
		Thread.sleep(2000);
		Driver.findElement(By.id("hrefAddNew")).click();
		Thread.sleep(2000);
		Driver.findElement(By.id("hrefAddNew")).click();
		Thread.sleep(2000);
		Driver.findElement(By.id("hrefAddNew")).click();
		Thread.sleep(2000);

		Driver.findElement(By.id("hlkSavePart")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("hlkSave")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
	}

	@Test
	public void UserProfile() throws Exception {
		// Read data from Excel
		// DEV
		// File src1 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUProfileDEV.xlsx");
		// Staging
		File src1 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUProfileSTG.xlsx");
		FileInputStream fis1 = new FileInputStream(src1);
		Workbook workbook = WorkbookFactory.create(fis1);
		// Sheet sh1 = workbook.getSheet("UserProfile");
		System.out.println("**********User Profile**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idUserProfile")).click();
		Thread.sleep(20000);
		Driver.findElement(By.id("hlkSavedTL")).click();
		Thread.sleep(5000);
		String Message1 = Driver.findElement(By.id("success")).getText();
		fis1.close();

		// DEV
		// File src2 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUProfileDEV.xlsx");
		// Staging
		File src2 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUProfileSTG.xlsx");
		FileOutputStream fis2 = new FileOutputStream(src2);
		Sheet sh2 = workbook.getSheet("UserProfile");

		if (Message1.equals("User Updated Successfully !")) {
			SheetMessage = "*****After Open Screen and Click on Save Button Data are save Successfully on User Profile Screen.*****";
			sh2.getRow(1).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "*****After Open Screen and Click on Save Button Data are not save Sucessfully on User Profile Screen.*****";
			sh2.getRow(1).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 1.)Login Type
		Boolean Ltype = Driver.findElement(By.id("txtLoginTypename")).isDisplayed();
		if (Ltype == true) {
			SheetMessage = "(1.) Login Type field is display on Screen.";
			sh2.getRow(2).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(1.) Login Type field is not display on Screen.";
			sh2.getRow(2).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		Boolean Ltype1 = Driver.findElement(By.id("txtLoginTypename")).getAttribute("readonly").equals("");
		if (Ltype1 == true) {
			SheetMessage = "(1.) Login Type field is Editable on Screen.";
			sh2.getRow(3).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(1.) Login Type field is not Editable on Screen.";
			sh2.getRow(3).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Ltype2 = Driver.findElement(By.id("txtLoginTypename")).getAttribute("value");
		if (Ltype2.isEmpty()) {
			SheetMessage = "(1.) Login Type is Empty.";
			sh2.getRow(4).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(1.) Login Type : " + Ltype2;
			sh2.getRow(4).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 2.)Reporting To
		Boolean RTo = Driver.findElement(By.id("txtReportingto")).isDisplayed();
		if (RTo == true) {
			SheetMessage = "(2.) Reporting To field is display on Screen.";
			sh2.getRow(5).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(2.) Reporting To field is not display on Screen.";
			sh2.getRow(5).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		Boolean RTo1 = Driver.findElement(By.id("txtReportingto")).getAttribute("readonly").equals("");
		if (RTo1 == true) {
			SheetMessage = "(2.) Reporting To field is Editable on Screen.";
			sh2.getRow(6).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(2.) Reporting To field is not Editable on Screen.";
			sh2.getRow(6).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String RTo2 = Driver.findElement(By.id("txtReportingto")).getAttribute("value");
		if (RTo2.equals(" ")) {
			SheetMessage = "(2.) Reporting To is Empty.";
			sh2.getRow(7).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(2.) Reporting To : " + RTo2;
			sh2.getRow(7).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 3.)Login ID
		Boolean LId = Driver.findElement(By.id("txtLoginId")).isDisplayed();
		if (LId == true) {
			SheetMessage = "(3.) Login ID field is display on Screen.";
			sh2.getRow(8).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(3.) Login ID field is not display on Screen.";
			sh2.getRow(8).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		Boolean LId1 = Driver.findElement(By.id("txtLoginId")).getAttribute("readonly").equals("");
		if (LId1 == true) {
			SheetMessage = "(3.) Login ID field is Editable on Screen.";
			sh2.getRow(9).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(3.) Login ID field is not Editable on Screen.";
			sh2.getRow(9).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String LId2 = Driver.findElement(By.id("txtLoginId")).getAttribute("value");
		if (LId2.isEmpty()) {
			SheetMessage = "(3.) Login ID is Empty.";
			sh2.getRow(10).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(3.) Login ID : " + LId2;
			sh2.getRow(10).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 6.)First Name
		Boolean FName = Driver.findElement(By.id("txtFirstName")).isDisplayed();
		if (FName == true) {
			SheetMessage = "(6.) First Name field is display on Screen.";
			sh2.getRow(11).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(6.) First Name field is not display on Screen.";
			sh2.getRow(11).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String FName1 = Driver.findElement(By.id("txtFirstName")).getAttribute("value");
		if (FName1.isEmpty()) {
			SheetMessage = "(6.) First Name text box is Empty.";
			sh2.getRow(12).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(6.) First Name : " + FName1;
			sh2.getRow(12).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 7.)Middle Name
		Boolean MName = Driver.findElement(By.id("txtMiddleName")).isDisplayed();
		if (MName == true) {
			SheetMessage = "(7.) Middle Name field is display on Screen.";
			sh2.getRow(13).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(7.) Middle Name field is not display on Screen.";
			sh2.getRow(13).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String MName1 = Driver.findElement(By.id("txtMiddleName")).getAttribute("value");
		if (MName1.isEmpty()) {
			SheetMessage = "(7.) Middle Name text box is Empty.";
			sh2.getRow(14).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(7.) Middle Name : " + MName1;
			sh2.getRow(14).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 8.)Last Name
		Boolean LName = Driver.findElement(By.id("txtLastName")).isDisplayed();
		if (LName == true) {
			SheetMessage = "(8.) Last Name field is display on Screen.";
			sh2.getRow(15).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(8.) Last Name field is not display on Screen.";
			sh2.getRow(15).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String LName1 = Driver.findElement(By.id("txtLastName")).getAttribute("value");
		if (LName1.isEmpty()) {
			SheetMessage = "(8.) Last Name text box is Empty.";
			sh2.getRow(16).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(8.) Last Name : " + LName1;
			sh2.getRow(16).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 9.)Title Name
		Boolean Title = Driver.findElement(By.id("txtTitle")).isDisplayed();
		if (Title == true) {
			SheetMessage = "(9.) Title field is display on Screen.";
			sh2.getRow(17).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(9.) Title field is not display on Screen.";
			sh2.getRow(17).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Title1 = Driver.findElement(By.id("txtTitle")).getAttribute("value");
		if (Title1.isEmpty()) {
			SheetMessage = "(9.) Title text box is Empty.";
			sh2.getRow(18).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(9.) Title : " + Title1;
			sh2.getRow(18).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 10.)Portal Type
		Boolean Ptype = Driver.findElement(By.id("txtPortaltypevalue")).isDisplayed();
		if (Ptype == true) {
			SheetMessage = "(10.) Portal Type field is display on Screen.";
			sh2.getRow(19).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(10.) Portal Type field is not display on Screen.";
			sh2.getRow(19).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		Boolean Ptype1 = Driver.findElement(By.id("txtPortaltypevalue")).getAttribute("readonly").equals("");
		if (Ptype1 == true) {
			SheetMessage = "(10.) Portal Type field is Editable on Screen.";
			sh2.getRow(20).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(10.) Portal Type field is not Editable on Screen.";
			sh2.getRow(20).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Ptype2 = Driver.findElement(By.id("txtPortaltypevalue")).getAttribute("value");
		if (Ptype2.isEmpty()) {
			SheetMessage = "(10.) Portal Type is Empty.";
			sh2.getRow(21).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(10.) Portal Type : " + Ptype2;
			sh2.getRow(21).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 11.)Description
		Boolean Desc = Driver.findElement(By.id("txtDescription")).isDisplayed();
		if (Desc == true) {
			SheetMessage = "(11.) Description field is display on Screen.";
			sh2.getRow(22).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(11.) Description field is not display on Screen.";
			sh2.getRow(22).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Desc1 = Driver.findElement(By.id("txtDescription")).getAttribute("value");
		if (Desc1.isEmpty()) {
			SheetMessage = "(11.) Description text box is Empty.";
			sh2.getRow(23).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(11.) Description : " + Desc1;
			sh2.getRow(23).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 12.)Password Last Set
		Boolean PLs = Driver.findElement(By.id("txtPwdlastSet")).isDisplayed();
		if (PLs == true) {
			SheetMessage = "(12.) Password Last Set field is display on Screen.";
			sh2.getRow(24).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(12.) Password Last Set field is not display on Screen.";
			sh2.getRow(24).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		Boolean PLs1 = Driver.findElement(By.id("txtPwdlastSet")).getAttribute("readonly").equals("");
		if (PLs1 == true) {
			SheetMessage = "(12.) Password Last Set field is Editable on Screen.";
			sh2.getRow(25).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(12.) Password Last Set field is not Editable on Screen.";
			sh2.getRow(25).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String PLs2 = Driver.findElement(By.id("txtPwdlastSet")).getAttribute("value");
		if (PLs2.isEmpty()) {
			SheetMessage = "(12.) Password Last Set is Empty.";
			sh2.getRow(26).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(12.) Portal Type : " + Ptype2;
			sh2.getRow(26).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 13.)Country = drpCountry
		Boolean COUN = Driver.findElement(By.id("drpCountry")).isDisplayed();
		if (COUN == true) {
			SheetMessage = "(13.) Country field is display on Screen.";
			sh2.getRow(27).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(13.) Country field is not display on Screen.";
			sh2.getRow(27).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String COUN1 = Driver.findElement(By.id("drpCountry")).getAttribute("value");
		if (COUN1.isEmpty()) {
			SheetMessage = "(13.) Country text box is Empty.";
			sh2.getRow(28).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(13.) Country : " + COUN1;
			sh2.getRow(28).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 14.)Zip/Postal Code = txtZipCode
		Boolean ZPCode = Driver.findElement(By.id("txtZipCode")).isDisplayed();
		if (ZPCode == true) {
			SheetMessage = "(14.) Zip/Postal Code field is display on Screen.";
			sh2.getRow(29).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(14.) Zip/Postal Code field is not display on Screen";
			sh2.getRow(29).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String ZPCode1 = Driver.findElement(By.id("txtZipCode")).getAttribute("value");
		if (ZPCode1.isEmpty()) {
			SheetMessage = "(14.) Zip/Postal Code text box is Empty.";
			sh2.getRow(30).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(14.) Zip/Postal Code : " + ZPCode1;
			sh2.getRow(30).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 15.)City = txtCity
		Boolean City = Driver.findElement(By.id("txtCity")).isDisplayed();
		if (City == true) {
			SheetMessage = "(15.) City field is display on Screen.";
			sh2.getRow(31).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(15.) City field is not display on Screen";
			sh2.getRow(31).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String City1 = Driver.findElement(By.id("txtCity")).getAttribute("value");
		if (City1.isEmpty()) {
			SheetMessage = "(15.) City text box is Empty.";
			sh2.getRow(32).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(15.) City : " + City1;
			sh2.getRow(32).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 16.)State = txtState
		Boolean Stat = Driver.findElement(By.id("txtState")).isDisplayed();
		if (Stat == true) {
			SheetMessage = "(16.) State field is display on Screen.";
			sh2.getRow(33).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(16.) State field is not display on Screen";
			sh2.getRow(33).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		Boolean Stat1 = Driver.findElement(By.id("txtState")).getAttribute("readonly").equals("");
		if (Stat1 == true) {
			SheetMessage = "(16.) State field is Editable on Screen.";
			sh2.getRow(34).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(16.) State field is not Editable on Screen";
			sh2.getRow(34).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Stat2 = Driver.findElement(By.id("txtState")).getAttribute("value");
		if (Stat2.isEmpty()) {
			SheetMessage = "(16.) State is Empty.";
			sh2.getRow(35).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(16.) State : " + Stat2;
			sh2.getRow(35).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 17.) Address Line 1 = txtAddr1
		Boolean Addr = Driver.findElement(By.id("txtAddr1")).isDisplayed();
		if (Addr == true) {
			SheetMessage = "(17.) Address Line 1 field is display on Screen.";
			sh2.getRow(36).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(17.) Address Line 1 field is not display on Screen";
			sh2.getRow(36).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Addr1 = Driver.findElement(By.id("txtAddr1")).getAttribute("value");
		if (Addr1.isEmpty()) {
			SheetMessage = "(17.) Address Line 1 text box is Empty.";
			sh2.getRow(37).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(17.) Address Line 1 : " + Addr1;
			sh2.getRow(37).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 18.) Dept/Suite = txtDept
		Boolean DSuite = Driver.findElement(By.id("txtDept")).isDisplayed();
		if (DSuite == true) {
			SheetMessage = "(18.) Dept/Suite field is display on Screen.";
			sh2.getRow(38).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(18.) Dept/Suite field is not display on Screen";
			sh2.getRow(38).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String DSuite1 = Driver.findElement(By.id("txtDept")).getAttribute("value");
		if (DSuite1.isEmpty()) {
			SheetMessage = "(18.) Dept/Suite text box is Empty.";
			sh2.getRow(39).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(18.) Dept/Suite : " + DSuite1;
			sh2.getRow(39).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 19.) Main Phone = txtMain
		Boolean MPhone = Driver.findElement(By.id("txtMain")).isDisplayed();
		if (MPhone == true) {
			SheetMessage = "(19.) Main Phone field is display on Screen.";
			sh2.getRow(40).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(19.) Main Phone field is not display on Screen";
			sh2.getRow(40).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String MPhone1 = Driver.findElement(By.id("txtMain")).getAttribute("value");
		if (MPhone1.isEmpty()) {
			SheetMessage = "(19.) Main Phone text box is Empty.";
			sh2.getRow(41).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(19.) Main Phone : " + MPhone1;
			sh2.getRow(41).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 20.) Extention = txtExt
		Boolean Ext = Driver.findElement(By.id("txtExt")).isDisplayed();
		if (Ext == true) {
			SheetMessage = "(20.) Extention field is display on Screen.";
			sh2.getRow(42).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(20.) Extention field is not display on Screen";
			sh2.getRow(42).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Ext1 = Driver.findElement(By.id("txtExt")).getAttribute("value");
		if (Ext1.isEmpty()) {
			SheetMessage = "(20.) Extention text box is Empty.";
			sh2.getRow(43).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(20.) Extention : " + Ext1;
			sh2.getRow(43).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 21.) Fax = txtFax
		Boolean Fax = Driver.findElement(By.id("txtFax")).isDisplayed();
		if (Fax == true) {
			SheetMessage = "(21.) Fax field is display on Screen.";
			sh2.getRow(44).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(21.) Fax field is not display on Screen";
			sh2.getRow(44).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Fax1 = Driver.findElement(By.id("txtFax")).getAttribute("value");
		if (Fax1.isEmpty()) {
			SheetMessage = "(21.) Fax text box is Empty.";
			sh2.getRow(45).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(21.) Fax : " + Fax1;
			sh2.getRow(45).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 22.) Email = txtEmail
		Boolean Email = Driver.findElement(By.id("txtEmail")).isDisplayed();
		if (Email == true) {
			SheetMessage = "(22.) Email field is display on Screen.";
			sh2.getRow(46).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(22.) Email field is not display on Screen";
			sh2.getRow(46).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Email1 = Driver.findElement(By.id("txtEmail")).getAttribute("value");
		if (Email1.isEmpty()) {
			SheetMessage = "(22.) Email text box is Empty.";
			sh2.getRow(47).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(22.) Email : " + Email1;
			sh2.getRow(47).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 23.) Work Phone = txtUserWorkphone
		Boolean WPhone = Driver.findElement(By.id("txtUserWorkphone")).isDisplayed();
		if (WPhone == true) {
			SheetMessage = "(23.) Work Phone field is display on Screen.";
			sh2.getRow(48).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(23.) Work Phone field is not display on Screen";
			sh2.getRow(48).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String WPhone1 = Driver.findElement(By.id("txtUserWorkphone")).getAttribute("value");
		if (WPhone1.isEmpty()) {
			SheetMessage = "(23.) Work Phone text box is Empty.";
			sh2.getRow(49).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(23.) Work Phone : " + WPhone1;
			sh2.getRow(49).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 24.) Extention = txtWorkphoneExt
		Boolean WExt = Driver.findElement(By.id("txtWorkphoneExt")).isDisplayed();
		if (WExt == true) {
			SheetMessage = "(24.) Work Extention field is display on Screen.";
			sh2.getRow(50).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(24.) Work Extention field is not display on Screen";
			sh2.getRow(50).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String WExt1 = Driver.findElement(By.id("txtWorkphoneExt")).getAttribute("value");
		if (WExt1.isEmpty()) {
			SheetMessage = "(24.) Work Extention text box is Empty.";
			sh2.getRow(51).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(24.) Work Extention : " + WExt1;
			sh2.getRow(51).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 25.) Cell Phone = txtCallphone
		Boolean CPhone = Driver.findElement(By.id("txtCallphone")).isDisplayed();
		if (CPhone == true) {
			SheetMessage = "(25.) Cell Phone field is display on Screen.";
			sh2.getRow(52).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(25.) Cell Phone field is not display on Screen";
			sh2.getRow(52).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String CPhone1 = Driver.findElement(By.id("txtCallphone")).getAttribute("value");
		if (CPhone1.isEmpty()) {
			SheetMessage = "(25.) Cell Phone text box is Empty.";
			sh2.getRow(53).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(25.) Cell Phone : " + CPhone1;
			sh2.getRow(53).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 26.) Home Phone = txtHomephone
		Boolean HPhone = Driver.findElement(By.id("txtHomephone")).isDisplayed();
		if (HPhone == true) {
			SheetMessage = "(26.) Home Phone field is display on Screen.";
			sh2.getRow(54).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(26.) Home Phone field is not display on Screen";
			sh2.getRow(54).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String HPhone1 = Driver.findElement(By.id("txtHomephone")).getAttribute("value");
		if (HPhone1.isEmpty()) {
			SheetMessage = "(26.) Home Phone text box is Empty.";
			sh2.getRow(55).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(26.) Home Phone : " + HPhone1;
			sh2.getRow(55).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 27.) Web Address = txtWebaddress
		Boolean WAddr = Driver.findElement(By.id("txtWebaddress")).isDisplayed();
		if (WAddr == true) {
			SheetMessage = "(27.) Web Address field is display on Screen.";
			sh2.getRow(56).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(27.) Web Address field is not display on Screen";
			sh2.getRow(56).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String WAddr1 = Driver.findElement(By.id("txtWebaddress")).getAttribute("value");
		if (WAddr1.isEmpty()) {
			SheetMessage = "(27.) Web Address text box is Empty.";
			sh2.getRow(57).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(27.) Web Address : " + WAddr1;
			sh2.getRow(57).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 28.) Security Question = txtSecQue
		Boolean SQue = Driver.findElement(By.id("txtSecQue")).isDisplayed();
		if (SQue == true) {
			SheetMessage = "(28.) Security Question field is display on Screen.";
			sh2.getRow(58).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(28.) Security Question field is not display on Screen";
			sh2.getRow(58).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String SQue1 = Driver.findElement(By.id("txtSecQue")).getAttribute("value");
		if (SQue1.isEmpty()) {
			SheetMessage = "(28.) Security Question text box is Empty.";
			sh2.getRow(59).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(28.) Security Question : " + SQue1;
			sh2.getRow(59).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// 29.) Response = txtSecAns
		Boolean SRespo = Driver.findElement(By.id("txtSecAns")).isDisplayed();
		if (SRespo == true) {
			SheetMessage = "(29.) Security Response field is display on Screen.";
			sh2.getRow(60).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(29.) Security Response field is not display on Screen";
			sh2.getRow(60).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String SRespo1 = Driver.findElement(By.id("txtSecAns")).getAttribute("value");
		if (SRespo1.isEmpty()) {
			SheetMessage = "(29.) Security Response text box is Empty.";
			sh2.getRow(61).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(29.) Security Response : " + SRespo1;
			sh2.getRow(61).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// for(int i=1;i<=61;i++)
		// {
		// SheetMessage = "=EXACT(A"+(i+1)+",B"+(i+1)+")";
		// sh2.getRow(i).createCell(2).setCellValue(SheetMessage);
		// }
		workbook.write(fis2);
		fis2.close();
		Thread.sleep(5000);
	}

	@Test
	public void Documents() throws Exception {
		System.out.println("**********Documents**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idCustDocuments")).click();
		Thread.sleep(20000);

		WebElement TogetRows = Driver.findElement(By.id("custdocuments"));
		List<WebElement> TotalRowsList = TogetRows.findElements(By.tagName("tr"));
		System.out.println("Total number of Rows in the table are : " + TotalRowsList.size());

		Thread.sleep(5000);

		WebElement ToGetColumns = Driver.findElement(By.id("custdocuments"));
		List<WebElement> TotalColsList = ToGetColumns.findElements(By.tagName("td"));
		System.out.println("Total Number of Columns in the table are: " + TotalColsList.size());

		Thread.sleep(5000);

		Driver.findElement(By.id("imgNew")).click();
		Thread.sleep(5000);

		Driver.findElement(By.id("hlkSaveShipPkg")).click();
		Thread.sleep(5000);

		String Message1 = Driver.findElement(By.id("idValidation")).getText();
		String today = CuDate();

		if (Message1.contains("Required")) {
			System.out.println(Message1);
			SheetMessage = "*****Documents Screen Validation are Completed !*****";
			System.out.println(SheetMessage);
		}

		Thread.sleep(5000);

		JavascriptExecutor js = (JavascriptExecutor) Driver;
		WebElement Element = Driver.findElement(By.id("btnUpload"));
		js.executeScript("arguments[0].scrollIntoView();", Element);

		Driver.findElement(By.id("txtDocName")).sendKeys("pdoshidoc-" + today);
		Thread.sleep(5000);

		String DocType = "Other";
		Select dropdown1 = new Select(Driver.findElement(By.id("drpDocType")));
		dropdown1.selectByVisibleText(DocType);

		Driver.findElement(By.id("txtDocDate")).sendKeys(today); // select today
		Thread.sleep(5000);

		Driver.findElement(By.id("txtRevision")).sendKeys("pdoshiRevision-" + today);
		Thread.sleep(5000);

		Driver.findElement(By.id("txtValidFrom")).click(); // click on calander
		Driver.findElement(By.id("txtValidFrom")).sendKeys(today); // select today
		Thread.sleep(5000);

		Driver.findElement(By.id("hlkSaveShipPkg")).click();
		Thread.sleep(5000);

		String Message2 = Driver.findElement(By.id("errorid")).getText();

		if (Message2.equals("Please select a file to upload")) {
			System.out.println(Message2);
			SheetMessage = "*****Documents Screen Validation are Completed !*****";
			System.out.println(SheetMessage);
		}

		Driver.findElement(By.id("file")).sendKeys(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\Consignee Address.xls");
		Thread.sleep(5000);
		Driver.findElement(By.id("btnUpload")).click();
		Thread.sleep(20000);

		if (Driver.findElement(By.id("successid")).isDisplayed() == true) {
			System.out.println(Driver.findElement(By.id("successid")).getText());
			SheetMessage = "*****Import Process Completed !*****";
			System.out.println(SheetMessage);
		}
		Thread.sleep(5000);

		js.executeScript("window.scrollBy(0,-250)");
		Thread.sleep(5000);

		Driver.findElement(By.id("hlkSaveShipPkg")).click();
		Thread.sleep(5000);

		try {
			String Message4 = Driver.findElement(By.id("successmsgid")).getText();
			Thread.sleep(5000);
			System.out.println("*****" + Message4 + "*****");
			Thread.sleep(5000);
		} catch (Exception e) {
			System.out.println("Sorry, There is no Successfully message.");
			Thread.sleep(5000);
		}

		try {
			String Message5 = Driver.findElement(By.id("errorid")).getText();
			Thread.sleep(5000);
			System.out.println("*****" + Message5 + "*****");
			Thread.sleep(5000);
		} catch (Exception e) {
			System.out.println("Greate, There is no Failure message.");
			Thread.sleep(5000);
		}

//		if(Message4.equals("Record Saved Successfully"))
//		{
//			System.out.println(Message4);
//			SheetMessage = "*****Record Saved Successfully.*****";
//			System.out.println(SheetMessage);
//		}
//		else if(Message5.contains("Document Type is already Exists."))
//		{
//			System.out.println(Message5);
//			SheetMessage = "*****Record is not Saved Successfully.*****";
//			System.out.println(SheetMessage);
//		}
	}

	@Test
	public void UserPrefrence0() throws Exception {
		// Read data from Excel
		// DEV
		// File src3 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUPreferenceDEV.xlsx");
		// STG
		File src3 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUPreferenceSTG.xlsx");
		FileInputStream fis3 = new FileInputStream(src3);
		Workbook workbook2 = WorkbookFactory.create(fis3);
		// Sheet sh3 = workbook2.getSheet("UserPrefrence");
		System.out.println("**********User Preference**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idUserPreferences")).click();
		Thread.sleep(20000);
		Driver.findElement(By.id("imgSave")).click();
		Thread.sleep(5000);
		String Message1 = Driver.findElement(By.id("success")).getText();
		fis3.close();

		// DEV
		// File src4 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUPreferenceDEV.xlsx");
		// Staging
		File src4 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipUPreferenceSTG.xlsx");
		FileOutputStream fis4 = new FileOutputStream(src4);
		Sheet sh4 = workbook2.getSheet("UserPrefrence");

		if (Message1.equals("Preference Updated successfully.")) {
			SheetMessage = "*****After Open Screen and Click on Save Button Data are save Successfully on User Preference Screen.*****";
			sh4.getRow(1).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "*****After Open Screen and Click on Save Button Data are not save Sucessfully on User Preference Screen.*****";
			sh4.getRow(1).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}

		// User Name : txtUsername
		// 1.)User Name
		Boolean UName = Driver.findElement(By.id("txtUsername")).isDisplayed();
		if (UName == true) {
			SheetMessage = "(1.) User Name field is display on Screen.";
			sh4.getRow(2).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(1.) User Name field is not display on Screen.";
			sh4.getRow(2).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Boolean UName1 =
		// Driver.findElement(By.id("txtUsername")).getAttribute("readonly").equals("");
		Boolean UName1 = Driver.findElement(By.id("txtUsername")).isEnabled();
		if (UName1 == true) {
			SheetMessage = "(1.) User Name field is Editable on Screen.";
			sh4.getRow(3).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(1.) User Name field is not Editable on Screen.";
			sh4.getRow(3).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String UName2 = Driver.findElement(By.id("txtUsername")).getAttribute("value");
		if (UName2.isEmpty()) {
			SheetMessage = "(1.) User Name is Empty.";
			sh4.getRow(4).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(1.) User Name : " + UName2;
			sh4.getRow(4).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Is Employee? : chkEmployee
		// 2.)IsEmployee
		// Currency : txtCurrency
		// 3.)Currency
		Boolean Curr = Driver.findElement(By.id("txtCurrency")).isDisplayed();
		if (Curr == true) {
			SheetMessage = "(3.) Currency field is display on Screen.";
			sh4.getRow(5).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(3.) Currency field is not display on Screen.";
			sh4.getRow(5).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Boolean Curr1 =
		// Driver.findElement(By.id("txtCurrency")).getAttribute("readonly").equals("");
		Boolean Curr1 = Driver.findElement(By.id("txtCurrency")).isEnabled();
		if (Curr1 == true) {
			SheetMessage = "(3.) Currency field is Editable on Screen.";
			sh4.getRow(6).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(3.) Currency field is not Editable on Screen.";
			sh4.getRow(6).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Curr2 = Driver.findElement(By.id("txtCurrency")).getAttribute("value");
		if (Curr2.isEmpty()) {
			SheetMessage = "(3.) Currency is Empty.";
			sh4.getRow(7).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(3.) Currency : " + Curr2;
			sh4.getRow(7).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Currency Symbol : txtCurrencySymbol
		// 4.)Currency Symbol
		Boolean CSym = Driver.findElement(By.id("txtCurrencySymbol")).isDisplayed();
		if (CSym == true) {
			SheetMessage = "(4.) Currency Symbol field is display on Screen.";
			sh4.getRow(8).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(4.) Currency Symbol field is not display on Screen.";
			sh4.getRow(8).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Boolean CSym1 =
		// Driver.findElement(By.id("txtCurrencySymbol")).getAttribute("readonly").equals("");
		Boolean CSym1 = Driver.findElement(By.id("txtCurrencySymbol")).isEnabled();
		if (CSym1 == true) {
			SheetMessage = "(4.) Currency Symbol field is Editable on Screen.";
			sh4.getRow(9).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(4.) Currency Symbol field is not Editable on Screen.";
			sh4.getRow(9).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String CSym2 = Driver.findElement(By.id("txtCurrencySymbol")).getAttribute("value");
		if (CSym2.isEmpty()) {
			SheetMessage = "(4.) Currency Symbol is Empty.";
			sh4.getRow(10).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(4.) Currency Symbol : " + CSym2;
			sh4.getRow(10).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Currency Seprator : ddlCurrencySeparator
		// 5.)Currency Seprator
		// Country : ddlCountry
		// 6.)Country
		Boolean COUN = Driver.findElement(By.id("ddlCountry")).isDisplayed();
		if (COUN == true) {
			SheetMessage = "(6.) Country field is display on Screen.";
			sh4.getRow(12).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(6.) Country field is not display on Screen.";
			sh4.getRow(12).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String COUN1 = Driver.findElement(By.id("ddlCountry")).getAttribute("value");
		if (COUN1.isEmpty()) {
			SheetMessage = "(6.) Country text box is Empty.";
			sh4.getRow(13).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(6.) Country : " + COUN1;
			sh4.getRow(13).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// State : ddlState
		// 7.)State
		Boolean State = Driver.findElement(By.id("ddlState")).isDisplayed();
		if (State == true) {
			SheetMessage = "(7.) State field is display on Screen.";
			sh4.getRow(14).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(7.) State field is not display on Screen.";
			sh4.getRow(14).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String State1 = Driver.findElement(By.id("ddlState")).getAttribute("value");
		if (State1.isEmpty()) {
			SheetMessage = "(7.) State text box is Empty.";
			sh4.getRow(15).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(7.) State : " + State1;
			sh4.getRow(15).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Culture : ddlCULTURE
		// 8.)Culture
		Boolean Culture = Driver.findElement(By.id("ddlCULTURE")).isDisplayed();
		if (Culture == true) {
			SheetMessage = "(8.) Culture field is display on Screen.";
			sh4.getRow(16).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(8.) Culture field is not display on Screen.";
			sh4.getRow(16).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String Culture1 = Driver.findElement(By.id("ddlCULTURE")).getAttribute("value");
		if (Culture1.isEmpty()) {
			SheetMessage = "(8.) Culture text box is Empty.";
			sh4.getRow(17).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(8.) Culture : " + Culture1;
			sh4.getRow(17).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Company : ddlCompany
		// 9.)Company
		Boolean COMP = Driver.findElement(By.id("ddlCompany")).isDisplayed();
		if (COMP == true) {
			SheetMessage = "(9.) Company field is display on Screen.";
			sh4.getRow(18).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(9.) Company field is not display on Screen.";
			sh4.getRow(18).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String COMP1 = Driver.findElement(By.id("ddlCompany")).getAttribute("value");
		if (COMP1.isEmpty()) {
			SheetMessage = "(9.) Company text box is Empty.";
			sh4.getRow(19).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(9.) Company : " + COMP1;
			sh4.getRow(19).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Time Zone : ddlTimeZone
		// 10.)Time Zone
		Boolean Tzone = Driver.findElement(By.id("ddlTimeZone")).isDisplayed();
		if (Tzone == true) {
			SheetMessage = "(10.) Time Zone field is display on Screen.";
			sh4.getRow(20).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(10.) Time Zone field is not display on Screen.";
			sh4.getRow(20).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		Select dropdown1 = new Select(Driver.findElement(By.id("ddlTimeZone")));
		WebElement dropdown2 = dropdown1.getFirstSelectedOption();
		String Tzone1 = dropdown2.getText().trim();
		if (Tzone1.isEmpty()) {
			SheetMessage = "(10.) Time Zone text box is Empty.";
			sh4.getRow(21).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(10.) Time Zone : " + Tzone1;
			sh4.getRow(21).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Date Format : ddlDateFormat
		// 11.)Date Format
		Boolean DFOR = Driver.findElement(By.id("ddlDateFormat")).isDisplayed();
		if (DFOR == true) {
			SheetMessage = "(11.) Date Format field is display on Screen.";
			sh4.getRow(22).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(11.) Date Format field is not display on Screen.";
			sh4.getRow(22).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String DFOR1 = Driver.findElement(By.id("ddlDateFormat")).getAttribute("value");
		if (DFOR1.isEmpty()) {
			SheetMessage = "(11.) Date Format text box is Empty.";
			sh4.getRow(23).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(11.) Date Format : " + DFOR1;
			sh4.getRow(23).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Date Seprator : ddlDateSeparator
		// 12.)Date Seprator
		Boolean DSEP = Driver.findElement(By.id("ddlDateSeparator")).isDisplayed();
		if (DSEP == true) {
			SheetMessage = "(12.) Date Seprator field is display on Screen.";
			sh4.getRow(24).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(12.) Date Seprator field is not display on Screen.";
			sh4.getRow(24).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String DSEP1 = Driver.findElement(By.id("ddlDateSeparator")).getAttribute("value");
		if (DSEP1.isEmpty()) {
			SheetMessage = "(12.) Date Seprator text box is Empty.";
			sh4.getRow(25).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(12.) Date Seprator : " + DSEP1;
			sh4.getRow(25).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Date/Time Format : ddlDateTimeFormat
		// 13.)Date/Time Format
		Boolean DTFOR = Driver.findElement(By.id("ddlDateTimeFormat")).isDisplayed();
		if (DTFOR == true) {
			SheetMessage = "(13.) Date/Time Format field is display on Screen.";
			sh4.getRow(26).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(13.) Date/Time Format field is not display on Screen.";
			sh4.getRow(26).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String DTFOR1 = Driver.findElement(By.id("ddlDateTimeFormat")).getAttribute("value");
		if (DTFOR1.isEmpty()) {
			SheetMessage = "(13.) Date/Time Format text box is Empty.";
			sh4.getRow(27).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(13.) Date/Time Format : " + DTFOR1;
			sh4.getRow(27).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Time Format : ddlTimeFormat
		// 14.)Time Format
		Boolean TFOR = Driver.findElement(By.id("ddlTimeFormat")).isDisplayed();
		if (TFOR == true) {
			SheetMessage = "(14.) Time Format field is display on Screen.";
			sh4.getRow(28).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(14.) Time Format field is not display on Screen.";
			sh4.getRow(28).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String TFOR1 = Driver.findElement(By.id("ddlTimeFormat")).getAttribute("value");
		if (TFOR1.isEmpty()) {
			SheetMessage = "(14.) Time Format text box is Empty.";
			sh4.getRow(29).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(14.) Time Format : " + TFOR1;
			sh4.getRow(29).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Valid To : ddlValidToMonth
		// 15.)Valid To
		Boolean VATO = Driver.findElement(By.id("ddlValidToMonth")).isDisplayed();
		if (VATO == true) {
			SheetMessage = "(15.) Valid To field is display on Screen.";
			sh4.getRow(30).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(15.) Valid To field is not display on Screen.";
			sh4.getRow(30).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String VATO1 = Driver.findElement(By.id("ddlValidToMonth")).getAttribute("value");
		if (VATO1.isEmpty()) {
			SheetMessage = "(15.) Valid To text box is Empty.";
			sh4.getRow(31).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(15.) Valid To : " + VATO1;
			sh4.getRow(31).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Grid Date/Time Format : ddlGridDateTimeFormat
		// 16.)Grid Date/Time Format
		Boolean GDTFOR = Driver.findElement(By.id("ddlGridDateTimeFormat")).isDisplayed();
		if (GDTFOR == true) {
			SheetMessage = "(16.) Grid Date/Time Format field is display on Screen.";
			sh4.getRow(32).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(16.) Grid Date/Time Format field is not display on Screen.";
			sh4.getRow(32).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String GDTFOR1 = Driver.findElement(By.id("ddlGridDateTimeFormat")).getAttribute("value");
		if (GDTFOR1.isEmpty()) {
			SheetMessage = "(16.) Grid Date/Time Format text box is Empty.";
			sh4.getRow(33).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(16.) Grid Date/Time Format : " + GDTFOR1;
			sh4.getRow(33).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// Memo DateTime Format : ddlMemoDateTimeFormat
		// 17.)Memo DateTime Format
		Boolean MDTFOR = Driver.findElement(By.id("ddlMemoDateTimeFormat")).isDisplayed();
		if (MDTFOR == true) {
			SheetMessage = "(17.) Memo DateTime Format field is display on Screen.";
			sh4.getRow(34).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(17.) Memo DateTime Format field is not display on Screen.";
			sh4.getRow(34).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String MDTFOR1 = Driver.findElement(By.id("ddlMemoDateTimeFormat")).getAttribute("value");
		if (MDTFOR1.isEmpty()) {
			SheetMessage = "(17.) Memo DateTime Format text box is Empty.";
			sh4.getRow(35).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(17.) Memo DateTime Format : " + MDTFOR1;
			sh4.getRow(35).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// LEN : ddlDefaultLEN
		// 18.)Length
		Boolean LEN = Driver.findElement(By.id("ddlDefaultLEN")).isDisplayed();
		if (LEN == true) {
			SheetMessage = "(18.) Length field is display on Screen.";
			sh4.getRow(36).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(18.) Length field is not display on Screen.";
			sh4.getRow(36).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String LEN1 = Driver.findElement(By.id("ddlDefaultLEN")).getAttribute("value");
		if (LEN1.isEmpty()) {
			SheetMessage = "(18.) Length text box is Empty.";
			sh4.getRow(37).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(18.) Length : " + LEN1;
			sh4.getRow(37).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// DRY : ddlDefaultDRY
		// 19.)DRY
		Boolean DRY = Driver.findElement(By.id("ddlDefaultDRY")).isDisplayed();
		if (DRY == true) {
			SheetMessage = "(19.) DRY field is display on Screen.";
			sh4.getRow(38).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(19.) DRY field is not display on Screen.";
			sh4.getRow(38).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String DRY1 = Driver.findElement(By.id("ddlDefaultDRY")).getAttribute("value");
		if (DRY1.isEmpty()) {
			SheetMessage = "(19.) DRY text box is Empty.";
			sh4.getRow(39).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(19.) DRY : " + DRY1;
			sh4.getRow(39).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// LEQ : ddlDefaultLIQ
		// 20.)Liquid
		Boolean LEQ = Driver.findElement(By.id("ddlDefaultLIQ")).isDisplayed();
		if (LEQ == true) {
			SheetMessage = "(20.) Liquid field is display on Screen.";
			sh4.getRow(40).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(20.) Liquid field is not display on Screen.";
			sh4.getRow(40).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String LEQ1 = Driver.findElement(By.id("ddlDefaultLIQ")).getAttribute("value");
		if (LEQ1.isEmpty()) {
			SheetMessage = "(20.) Liquid text box is Empty.";
			sh4.getRow(41).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(20.) Liquid : " + LEQ1;
			sh4.getRow(41).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// WGT : ddlDefaultWGT
		// 21.)Weight
		Boolean WGT = Driver.findElement(By.id("ddlDefaultWGT")).isDisplayed();
		if (WGT == true) {
			SheetMessage = "(21.) Weight field is display on Screen.";
			sh4.getRow(42).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(21.) Weight field is not display on Screen.";
			sh4.getRow(42).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String WGT1 = Driver.findElement(By.id("ddlDefaultWGT")).getAttribute("value");
		if (WGT1.isEmpty()) {
			SheetMessage = "(21.) Weight text box is Empty.";
			sh4.getRow(43).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(21.) Weight : " + WGT1;
			sh4.getRow(43).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		// No.of Pages : txtNoOfPages
		// 22.)No.of Pages
		Boolean NOPage = Driver.findElement(By.id("txtNoOfPages")).isDisplayed();
		if (NOPage == true) {
			SheetMessage = "(22.) No.of Pages field is display on Screen.";
			sh4.getRow(44).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(22.) No.of Pages field is not display on Screen.";
			sh4.getRow(44).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		String NOPage1 = Driver.findElement(By.id("txtNoOfPages")).getAttribute("value");
		if (NOPage1.isEmpty()) {
			SheetMessage = "(22.) No.of Pages text box is Empty.";
			sh4.getRow(45).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		} else {
			SheetMessage = "(22.) No.of Pages : " + NOPage1;
			sh4.getRow(45).createCell(1).setCellValue(SheetMessage);
			System.out.println(SheetMessage);
		}
		workbook2.write(fis4);
		fis4.close();
	}

	@Test
	public void UserPrefrence1() throws Exception {
		Driver.findElement(By.id("imgLogo")).click();
		Thread.sleep(15000);
		String ActiveOrderData = Driver.findElement(By.id("ActiveOrderGrd")).getText();
		System.out.println(ActiveOrderData + "\n\n\n\n");
		String list[] = ActiveOrderData.split(" ");
		String list1[] = list[1].split("\n");
		System.out.println("First :- " + list[0]);
		System.out.println("Second :- " + list[1]);
		System.out.println("Third :- " + list[2]);
		System.out.println("Fourth :- " + list[3]);
		System.out.println("Fifth :- " + list[4]);
		System.out.println("\n\n\n\n\n");
		System.out.println("First :- " + list1[0]);
		System.out.println("Second :- " + list1[1]);
		// System.out.println("Third :- " + list1 [2]);
		// System.out.println("Fourth :- " + list1 [3]);
		// System.out.println("Fifth :- " + list1 [4]);
		Thread.sleep(5000);

		String PUID = "PickupId_N" + list1[0];
		Driver.findElement(By.id(PUID)).click();
		Thread.sleep(15000);

		getscreenshot("UserPreference1-01");
		Thread.sleep(5000);

		System.out.println("**********User Preference**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idUserPreferences")).click();
		Thread.sleep(20000);

		Select dropdown1 = new Select(Driver.findElement(By.id("ddlTimeZone")));
		WebElement dropdown2 = dropdown1.getFirstSelectedOption();
		int a = dropdown2.getText().length();
		String Tzone1 = dropdown2.getText().trim();
		int b = Tzone1.length();

		String Tzone2 = "America/Chicago (UTC-6)";// -CST
		String Tzone3 = "America/New_York (UTC-5)";// -EST
		String Tzone4 = "America/Phoenix (UTC-7)";// -MST
		String Tzone5 = "America/Los_Angeles (UTC-8)";// -PST

		System.out.println("Length ====> " + a);
		System.out.println("Length ====> " + b);

		if (Tzone1.equals(Tzone2)) {
			dropdown1.selectByVisibleText("America/Los_Angeles (UTC-8)");
			Thread.sleep(5000);
			Driver.findElement(By.id("imgSave")).click();
			Thread.sleep(5000);
		} else if (Tzone1.equals(Tzone3)) {
			dropdown1.selectByVisibleText("America/Phoenix (UTC-7)");
			Thread.sleep(5000);
			Driver.findElement(By.id("imgSave")).click();
			Thread.sleep(5000);
		} else if (Tzone1.equals(Tzone4)) {
			dropdown1.selectByVisibleText("America/New_York (UTC-5)");
			Thread.sleep(5000);
			Driver.findElement(By.id("imgSave")).click();
			Thread.sleep(5000);
		} else if (Tzone1.equals(Tzone5)) {
			dropdown1.selectByVisibleText("America/Chicago (UTC-6)");
			Thread.sleep(5000);
			Driver.findElement(By.id("imgSave")).click();
			Thread.sleep(5000);
		} else {
			System.out.println("User Preference Screen have " + Tzone1 + "Time Zone Instead of " + Tzone2 + ", "
					+ Tzone3 + ", " + Tzone4 + ", " + Tzone5 + ".....");
		}

		Thread.sleep(5000);
		try {
			System.out.println(Driver.findElement(By.className("success-messages")).getText());
			Thread.sleep(5000);
		} catch (Exception e) {
			System.out.println("There is no Success Message Display.");
		}

		Driver.findElement(By.id("lbltotalorders")).click();
		Thread.sleep(20000);

		System.out.println("**********Logout**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("hrefLogout")).click();
		Thread.sleep(5000);

		System.out.println("**********Login Sucessfully**********");
		// ********************User Name and Password***********************
		// DEV
		// Driver.get("http://10.20.104.122:8075/login");
		// Staging
		Driver.get("http://stagingns.nglog.com/");
		// Pre-Production
		// Driver.get("http://192.168.11.82:8074/");

		Driver.findElement(By.id("inputUsername")).clear();
		// DEV
		// Driver.findElement(By.id("inputUsername")).sendKeys("95008401");
		// Staging
		Driver.findElement(By.id("inputUsername")).sendKeys("95002401");
		// Pre-Production
		// driver.findElement(By.id("inputUsername")).sendKeys("10327201");

		Driver.findElement(By.id("inputPassword")).clear();
		// DEV AND Staging
		Driver.findElement(By.id("inputPassword")).sendKeys("pdoshi");
		// Pre-Production
		// driver.findElement(By.id("inputPassword")).sendKeys("password");
		Driver.findElement(By.id("btnSignIn")).click();
		Thread.sleep(20000);

		System.out.println("**********Net Ship Information Popup**********");
		try {
			if (Driver.findElement(By.id("btnDismiss")).isDisplayed() == true) {
				getscreenshot("NetShipInfoPopup");
				Thread.sleep(5000);
				Driver.findElement(By.id("btnDismiss")).click();
				Thread.sleep(5000);
				System.out.println("Net Ship Info Pop up is display.");
			}
		} catch (Exception e) {
			System.out.println("Net Ship Info Pop up is not display.");
		}
		Thread.sleep(20000);

		Driver.findElement(By.id(PUID)).click();
		Thread.sleep(15000);

		getscreenshot("UserPreference1-02");
		Thread.sleep(5000);
	}

	@Test
	public void OrderPrefrence0() throws Exception {
		// Read data from Excel
		// DEV
		// File src0 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipOrderPrefeenceDEV.xlsx");
		// Staging
		File src0 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipOrderPrefeenceSTG.xlsx");
		FileInputStream fis0 = new FileInputStream(src0);
		Workbook workbook = WorkbookFactory.create(fis0);
		Sheet sh0 = workbook.getSheet("OrderPreference0");
		int rcount = sh0.getLastRowNum();
		DataFormatter formatter = new DataFormatter();
		System.out.println("**********Order Preference**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.id("idOrderPreferences")).click();
		Thread.sleep(20000);

		System.out.println("Total Row ==> " + rcount);

		// Start from second row in Excel
		for (int i = 2; i <= rcount; i++) {
			Thread.sleep(5000);
			Driver.findElement(By.id("saveOrderPref")).click();
			Thread.sleep(5000);
			String Message1 = Driver.findElement(By.id("successmsgid")).getText();

			String CustomerName = formatter.formatCellValue(sh0.getRow(i).getCell(1));

			Select dropdown1 = new Select(Driver.findElement(By.id("drpClient")));
			dropdown1.selectByVisibleText(CustomerName);
			Thread.sleep(15000);
			Driver.findElement(By.id("saveOrderPref")).click();
			Thread.sleep(15000);

			String Message2 = Driver.findElement(By.id("successmsgid")).getText();

			if (Message1.equals("Order Preference Updated Successfully!")) {
				SheetMessage = "*****After Open Screen and Click on Save Button Data are save Successfully on Order Preference Screen.*****";
				System.out.println(SheetMessage);
			} else {
				SheetMessage = "*****After Open Screen and Click on Save Button Data are not save Sucessfully on Order Preference Screen.*****";
				System.out.println(SheetMessage);
			}

			if (Message2.equals("Order Preference Updated Successfully!")) {
				SheetMessage = "*****After Select Customer from Combo and Click on Save Button Data are save Successfully on Order Preference Screen.*****";
				System.out.println(SheetMessage);
			} else {
				SheetMessage = "*****After Select Customer from Combo and Click on Save Button Data are not save Sucessfully on Order Preference Screen.*****";
				System.out.println(SheetMessage);
			}
			Thread.sleep(5000);
			String find1 = Driver.findElement(By.id("txtAddress")).getAttribute("value");
			System.out.println(find1);
			if (find1.isEmpty()) {
				Select CustDropDown = new Select(Driver.findElement(By.id("drpClient")));
				CustDropDown.selectByVisibleText(CustomerName);
				Thread.sleep(15000);

				Driver.findElement(By.id("txtTime")).clear();
				Driver.findElement(By.id("txtTime")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(26)));
				Driver.findElement(By.id("chkReadyNow")).click();
				Driver.findElement(By.id("txtCompany")).clear();
				Driver.findElement(By.id("txtCompany")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(2)));
				Driver.findElement(By.id("txtShipper")).clear();
				Driver.findElement(By.id("txtShipper")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(3)));
				Driver.findElement(By.id("txtPhone")).clear();
				Driver.findElement(By.id("txtPhone")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(4)));
				Driver.findElement(By.id("txtphone1")).clear();
				Driver.findElement(By.id("txtphone1")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(5)));
				Driver.findElement(By.id("txtAddress")).clear();
				Driver.findElement(By.id("txtAddress")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(6)));
				Driver.findElement(By.id("txtDeptSuite")).clear();
				Driver.findElement(By.id("txtDeptSuite")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(7)));

				Thread.sleep(5000);
				Select CounDropDown = new Select(Driver.findElement(By.id("drpCountry")));
				CounDropDown.selectByVisibleText(formatter.formatCellValue(sh0.getRow(i).getCell(8)));
				Thread.sleep(5000);

				Driver.findElement(By.id("txtPUZipCode")).clear();
				Driver.findElement(By.id("txtPUZipCode")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(9)));
				Thread.sleep(5000);
				Driver.findElement(By.id("txtPUZipCode")).sendKeys(Keys.ENTER);
				Thread.sleep(5000);
				Select PUCityDropDown = new Select(Driver.findElement(By.id("txtCity")));
				WebElement option1 = PUCityDropDown.getFirstSelectedOption();
				String City1 = option1.getText();
				String State1 = Driver.findElement(By.id("txtState")).getAttribute("value");
				if (City1.isEmpty() || State1.isEmpty()) {
					System.out.print("Pickup City and State BOTH are not bind Proper.");
				} else {
					System.out.print("Pickup City and State BOTH are bind Proper.");
				}
				Driver.findElement(By.id("txtInstruction")).clear();
				Driver.findElement(By.id("txtInstruction"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(12)));
				Driver.findElement(By.id("txtCompanydel")).clear();
				Driver.findElement(By.id("txtCompanydel"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(13)));
				Driver.findElement(By.id("txtAttention")).clear();
				Driver.findElement(By.id("txtAttention"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(14)));
				Driver.findElement(By.id("txtPhonedel")).clear();
				Driver.findElement(By.id("txtPhonedel")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(15)));
				Driver.findElement(By.id("txtphonedel1")).clear();
				Driver.findElement(By.id("txtphonedel1"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(16)));
				Driver.findElement(By.id("txtAddressdel")).clear();
				Driver.findElement(By.id("txtAddressdel"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(17)));
				Driver.findElement(By.id("txtDeptSuitedel")).clear();
				Driver.findElement(By.id("txtDeptSuitedel"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(18)));

				Thread.sleep(5000);
				Select DLCounDropDown = new Select(Driver.findElement(By.id("drpDLCountry")));
				DLCounDropDown.selectByVisibleText(formatter.formatCellValue(sh0.getRow(i).getCell(19)));
				Thread.sleep(5000);

				Driver.findElement(By.id("txtDLZipCode")).clear();
				Driver.findElement(By.id("txtDLZipCode"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(20)));
				Thread.sleep(5000);
				Driver.findElement(By.id("txtDLZipCode")).sendKeys(Keys.ENTER);
				Thread.sleep(5000);
				Select DLCityDropDown = new Select(Driver.findElement(By.id("txtCity")));
				WebElement option2 = DLCityDropDown.getFirstSelectedOption();
				String City2 = option2.getText();
				String State2 = Driver.findElement(By.id("txtState")).getAttribute("value");
				if (City2.isEmpty() || State2.isEmpty()) {
					System.out.print("\nDeliver City and State BOTH are not bind Proper.");
				} else {
					System.out.print("\nDeliver City and State BOTH are bind Proper.");
				}
				Driver.findElement(By.id("txtInstructiondel")).clear();
				Driver.findElement(By.id("txtInstructiondel"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(23)));
				Driver.findElement(By.id("txtEmail")).clear();
				Driver.findElement(By.id("txtEmail")).sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(28)));
				Driver.findElement(By.id("chkOrderReceived")).click();
				Driver.findElement(By.id("chkAlDrop")).click();
				Driver.findElement(By.id("chkQdtChange")).click();
				Driver.findElement(By.id("chkOrderProcess")).click();
				Driver.findElement(By.id("chkAlRecover")).click();
				Driver.findElement(By.id("chkPickedUp")).click();
				Driver.findElement(By.id("chkDeliveredup")).click();
				Driver.findElement(By.id("txtConsgEmail")).clear();
				Driver.findElement(By.id("txtConsgEmail"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(29)));
				Driver.findElement(By.id("chkOrderReceivednot")).click();
				Driver.findElement(By.id("chkAlDropnot")).click();
				Driver.findElement(By.id("chkQdtChangenot")).click();
				Driver.findElement(By.id("chkOrderProcessnot")).click();
				Driver.findElement(By.id("chkAlRecovernot")).click();
				Driver.findElement(By.id("chkPickedUpnot")).click();
				Driver.findElement(By.id("chkDelivered")).click();

				Thread.sleep(5000);
				Select ModeDropDown = new Select(Driver.findElement(By.id("drpMode")));
				ModeDropDown.selectByVisibleText(formatter.formatCellValue(sh0.getRow(i).getCell(27)));
				Thread.sleep(5000);

				Driver.findElement(By.id("chkyes")).click();
				Thread.sleep(5000);

				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_PAGE_DOWN);
				Thread.sleep(5000);
				robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
				Thread.sleep(5000);

				Driver.findElement(By.xpath(
						"/html/body/div[3]/section/div[2]/div[1]/div/div[2]/div[2]/div[3]/div[7]/div[1]/div/div[2]/label"))
						.click();
				Thread.sleep(5000);
				Driver.findElement(By.id("chkyesadd")).click();
				Thread.sleep(5000);
				Driver.findElement(By.id("txtReference")).clear();
				Thread.sleep(5000);
				Driver.findElement(By.id("txtReference"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(23)));

				Thread.sleep(5000);
				Select DefCustDropDown = new Select(Driver.findElement(By.id("drpClientOrderPref")));
				DefCustDropDown.selectByVisibleText("No Default");
				Thread.sleep(5000);

				Driver.findElement(By.id("txtCommodity")).clear();
				Driver.findElement(By.id("txtCommodity"))
						.sendKeys(formatter.formatCellValue(sh0.getRow(i).getCell(24)));

				Select ServiceDrop = new Select(Driver.findElement(By.id("cmbdefualtservice")));
				ServiceDrop.selectByVisibleText(formatter.formatCellValue(sh0.getRow(i).getCell(30)));
				Thread.sleep(5000);

				Driver.findElement(By.id("lblOrderPrefService")).click();
				Thread.sleep(5000);

				robot.keyPress(KeyEvent.VK_PAGE_UP);
				Thread.sleep(5000);
				robot.keyRelease(KeyEvent.VK_PAGE_UP);
				Thread.sleep(5000);

				Driver.findElement(By.id("saveOrderPref")).click();
				Thread.sleep(15000);

				String DataSave1 = Driver.findElement(By.id("successmsgid")).getText();

				if (DataSave1.equals("Order Preference Updated Successfully!")) {
					SheetMessage = "*****Data are Save Successfully for " + CustomerName + " Customer.*****";
					System.out.println(SheetMessage);
				} else {
					SheetMessage = "*****Data are not Save Successfully for " + CustomerName + " Customer.*****";
					System.out.println(SheetMessage);
					System.out.println("After Click on Save This Message is display '*****" + DataSave1 + "*****'");
				}
			} else {
				System.out.println("(" + CustomerName + ")" + " :- " + "This Customer have Order Preference Detail.");
			}
		}
	}

	@Test
	public void ContactUS() throws Exception {
		// Read data from Excel
		// DEV
		// File src3 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipContactUSDEV.xlsx");
		// Staging
		File src3 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipContactUSSTG.xlsx");
		FileInputStream fis3 = new FileInputStream(src3);
		Workbook workbook2 = WorkbookFactory.create(fis3);
		// Sheet sh3 = workbook2.getSheet("ContactUS");
		System.out.println("**********Contact US**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath("//*[@id=\"headerDiv\"]/div/div[3]/ul/li[4]/a")).click();
		Thread.sleep(15000);
		fis3.close();

		// DEV
		// File src4 = new
		// File("D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipContactUSDEV.xlsx");
		// Staging
		File src4 = new File(
				"D:\\Netlink_AJ\\TestAutomation\\ConnectDetailTesting\\ConnectEclipse\\NetShipAllScreen\\DataFiles\\NetShipContactUSSTG.xlsx");
		FileOutputStream fis4 = new FileOutputStream(src4);
		Sheet sh4 = workbook2.getSheet("ContactUS");

		String Data0 = Driver.findElement(By.xpath("/html/body/div[3]/section/div[2]/div[1]/div/div[2]")).getText();

		SheetMessage = Data0;
		if (SheetMessage.equals(Data0)) {
			getscreenshot("ContactUS01");
			sh4.getRow(1).createCell(1).setCellValue(SheetMessage);
			System.out.println("====================================================================================");
			System.out.println(SheetMessage);
			System.out.println("====================================================================================");
			workbook2.write(fis4);
			fis4.close();
		} else {
			System.out.println("====================================================================================");
			System.out.println("Contact US screen data are not same.??????????");
			System.out.println("====================================================================================");
			System.out.println(SheetMessage);
			System.out.println("====================================================================================");
			System.out.println(Data0);
			System.out.println("====================================================================================");
		}
	}

	@Test
	public void NewsANDAnnouncements() throws Exception {
		System.out.println("**********News & Announcements**********");
		Thread.sleep(5000);
		Driver.findElement(By.id("divUsername")).click();
		Thread.sleep(5000);
		Driver.findElement(By.xpath("//*[@id=\"headerDiv\"]/div/div[3]/ul/li[5]/a")).click();
		Thread.sleep(15000);
		System.out.println("News & Announcements Screen Open Susseccfully.");
		getscreenshot("News & Announcements01");
	}

}