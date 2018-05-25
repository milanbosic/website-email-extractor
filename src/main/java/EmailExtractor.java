import org.apache.commons.validator.routines.UrlValidator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.openqa.selenium.By;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EmailExtractor {
    private String url;
    public static WebDriver driver;

    private EmailExtractor(String url) {
        this.url = url;
    }


    private HashSet<String> getEmails() {

        HashSet<String> urls = new HashSet<>();
        HashSet<String> emails = new HashSet<>();
        UrlValidator urlValidator = new UrlValidator();

        /*RegEx pattern used for finding emails*/
        Pattern p = Pattern.compile("(?!\\S*\\.(?:jpg|png|gif|bmp)(?:[\\s\\n\\r]|$))\\b[\\w.%-]+@[-.\\w]+\\.[A-Za-z]{2,4}\\b");

        /*open a connection to the website and fetch its HTML*/
        try {

            WebDriver driver = fetchHTMLDriver(url);
            emails = checkForEmails(driver, p);
            String sourceHTML = driver.getPageSource();

            /*if body length or header size is smaller than the given number, HTML probably failed to be fetched
             * or the website is dead*/
            if (sourceHTML != null) {

                if (sourceHTML.length() < 20 || !sourceHTML.toLowerCase().contains("frame")) {
                    System.out.println("Probably invalid HTML");
                    emails.add("dead URL");
                }
            }

            /*if no emails have been found in the base URL, check the links that contain keywords,
             * such as 'contact' or' 'about'
             */
            String linktemp = driver.getCurrentUrl().toLowerCase();
            if (linktemp.contains("domain") || linktemp.contains("newvcorp") || linktemp.contains("afternic") || linktemp.contains("undeveloped")
                    || linktemp.contains("godaddy") || linktemp.contains("hostgator") || linktemp.contains("bluehost") || linktemp.contains("uniregistry")){
                emails.add("dead URL");
            }

            if (emails.isEmpty()) {
                List<WebElement> allLinks = driver.findElements(By.tagName("a"));

                try {
                    for (WebElement link : allLinks) {
                        String currentText = link.getText();
                        if (currentText == null) {
                            currentText = "empty";
                        }
                        String currentLink = link.getAttribute("href");
                        if (currentLink == null) {
                            currentLink = "empty";
                        }
                        String temp = currentText.toLowerCase() + currentLink.toLowerCase();
                        if (temp.contains("contact") || temp.contains("location") || temp.contains("about") || temp.contains("connect") || temp.contains("welcome")) {

                            if (urlValidator.isValid(currentLink)) {
                                System.out.println("Found valid URL, adding: " + currentLink);
                                urls.add(currentLink);
                            } else {
                                System.out.println("Found invalid URL, need to concatenate: " + currentLink);
                                Connection.Response response = Jsoup.connect(url).followRedirects(true).execute();
                                System.out.println("Response url from jsoup: " + response.url());

                                String responseURL = driver.getCurrentUrl();
                                System.out.println("Response url from selenium: " + responseURL);
                                HashSet<String> allurls = new HashSet<>();

                                String concatenatedURL = concatenateURL(responseURL, currentLink);
                                if (urlValidator.isValid(concatenatedURL)) {
                                    allurls.add(concatenatedURL);
                                }

                                if (responseURL.contains("index")) {
                                    System.out.println("Base URL contains 'index', removing...");
                                    responseURL = responseURL.split("index")[0];
                                    concatenatedURL = concatenateURL(responseURL, currentLink);
                                    allurls.add(concatenatedURL);
                                } else if (responseURL.contains("home")) {
                                    System.out.println("Base URL contains 'home', removing...");
                                    responseURL = responseURL.split("home")[0];
                                    concatenatedURL = concatenateURL(responseURL, currentLink);
                                    allurls.add(concatenatedURL);
                                }

                                for (String str : allurls) {
                                    System.out.println("Concatenated URL in HashSet: " + str);
                                    if (urlValidator.isValid(str)) {
                                        urls.add(str);
                                    }

                                }
                            }

                        }
                    }

                } catch (StaleElementReferenceException e) {
                    e.printStackTrace();
                }


            }

            if (!urls.isEmpty()) {
                for (String url : urls) {
                    System.out.println("Current url being checked: " + url);
                    driver = fetchHTMLDriver(url);

                    HashSet<String> temp = checkForEmails(driver, p);

                    if (!temp.isEmpty()) {
                        emails.addAll(temp);
                    }

                }
            }


        } catch (IOException e) {
            //e.printStackTrace();
            System.out.println("URL appears to be dead");
            emails.add("dead URL");
        }

        if (!emails.isEmpty()) {
            System.out.println("Final emails: ");
            for (String email : emails) {
                System.out.println(email);
            }

        } else {
            System.out.println("No emails found!");
            emails.add("no");
        }
        return emails;

    }

    private WebDriver fetchHTMLDriver(String url) {

        driver.get(url);
        return driver;

    }

    private HashSet<String> checkForEmails(WebDriver driver, Pattern p) {
        HashSet<String> emails = new HashSet<>();
        Matcher matcher;

        String stringMatcher = "empty";

        if (driver.getPageSource() != null) {
            stringMatcher = driver.getPageSource();
        }

        matcher = p.matcher(stringMatcher);

        while (matcher.find()) {
            System.out.println("Mail found in base URL: " + matcher.group());
            emails.add(matcher.group());

        }

        return emails;

    }

    /*concatenate urls that are not valid, eg. '/contact.html' with the base url*/
    private String concatenateURL(String responseURL, String currentLink) {
        String concatenatedURL;
        if (currentLink.startsWith("/") && !responseURL.endsWith("/")) {
            concatenatedURL = responseURL + currentLink;
        } else if (currentLink.startsWith("/") && responseURL.endsWith("/")) {
            StringBuilder sb = new StringBuilder(currentLink);
            sb.deleteCharAt(0);
            concatenatedURL = responseURL + sb.toString();
        } else if (!currentLink.startsWith("/") && responseURL.endsWith("/")) {
            concatenatedURL = responseURL + currentLink;
        } else {
            concatenatedURL = responseURL + "/" + currentLink;
        }
        return concatenatedURL;
    }

    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("headless");
        options.addArguments("window-size=1200x600");

        driver = new ChromeDriver(options);

        try {
            /*open input stream and access the .xlsx file and its sheet*/
            FileInputStream fileInputStream = new FileInputStream(new File("C:/Users/milan/Desktop/test.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);

            Scanner scanner = new Scanner(System.in);
            System.out.println("Enter column number with URLs (A=0, B=1, etc.): ");
            int columnURLs = scanner.nextInt();
            System.out.println("Enter column number to write emails to (A=0, B=1, etc.): ");
            int columnEmails = scanner.nextInt();
            System.out.println("Enter the starting row: ");
            int startingRow = scanner.nextInt() - 1;
            System.out.println("Enter the number of rows to process: ");
            int numberOfRows = scanner.nextInt();
            int adjustedNumberOfRows = startingRow + numberOfRows;

            /*iterate through the rows with urls*/
            for (int i = startingRow; i < adjustedNumberOfRows; i++) {
                String url = sheet.getRow(i).getCell(columnURLs).toString();
                EmailExtractor emailExtractor = new EmailExtractor(url);
                int rownum = i + 1;
                System.out.println("Row number: " + rownum + " Base URL: " + url);
                HashSet<String> emails = emailExtractor.getEmails();

                if (!emails.isEmpty()) {
                    try {
                        FileOutputStream fileOutputStream = new FileOutputStream(new File("C:/Users/milan/Desktop/test.xlsx"));
                        Cell cell = sheet.getRow(i).getCell(columnEmails);
                        List<String> emailList = new ArrayList<>(emails);
                        if (cell == null) {
                            cell = sheet.getRow(i).createCell(columnEmails);
                        }

                        /*write the preferred email to cell, in this case
                        ones that contain the words 'info' or 'contact'*/
                        boolean foundPreferredEmails = false;
                        for (String s : emailList) {
                            if (s.contains("info")) {
                                cell.setCellValue(s);
                                foundPreferredEmails = true;
                                break;
                            } else if (s.contains("contact")) {
                                cell.setCellValue(s);
                                foundPreferredEmails = true;
                                break;
                            }
                        }

                        /*if no emails that match the criteria had been found, write the
                        first email in the array as the default*/
                        if (!foundPreferredEmails) {
                            cell.setCellValue(emailList.get(0));
                        }
                        /*write to workbook and close the stream*/
                        workbook.write(fileOutputStream);
                        fileOutputStream.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
