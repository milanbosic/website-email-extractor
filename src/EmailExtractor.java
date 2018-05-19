import org.apache.commons.validator.routines.UrlValidator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EmailExtractor {
    private String url;

    private EmailExtractor(String url) {
        this.url = url;
    }


    private HashSet<String> getEmails() {

        HashSet<String> urls = new HashSet<>();
        HashSet<String> emails = new HashSet<>();
        UrlValidator urlValidator = new UrlValidator();

        /*RegEx pattern used for finding emails*/
        Pattern p = Pattern.compile("\\b[\\w.%-]+@[-.\\w]+\\.[A-Za-z]{2,4}\\b");

        /*open a connection to the website and fetch its HTML*/
        try {

            Document doc = fetchHTMLDocument(url);
            emails = checkForEmails(doc, p);

            /*if body length or header size is smaller than the given number, HTML probably failed to be fetched
             * or the website is dead*/
            if (doc.body() != null && doc.head() != null) {

                if (doc.body().toString().length() < 1000 && doc.head().toString().length() < 500) {
                    System.out.println("Probably invalid HTML");
                    emails.add("dead URL");
                }
            }

            /*if no emails have been found in the base URL, check the links that contain keywords,
             * such as 'contact' or' 'about'
             */
            if (emails.isEmpty()) {
                Elements links = doc.select("a[href]");

                for (Element link : links) {
                    String currentText = link.text();
                    String currentLink = link.attr("href");
                    String temp = currentText.toLowerCase() + currentLink.toLowerCase();
                    if (temp.contains("contact") || temp.contains("location") || temp.contains("about")) {

                        if (urlValidator.isValid(currentLink)) {
                            System.out.println("Found valid URL, adding: " + currentLink);
                            urls.add(currentLink);
                        } else {
                            System.out.println("Found invalid URL, need to concatenate: " + currentLink);
                            Connection.Response response = Jsoup.connect(url).followRedirects(true).execute();
                            System.out.println("Response url: " + response.url());

                            String responseURL = response.url().toString();
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
            }

            if (!urls.isEmpty()) {
                for (String url : urls) {
                    System.out.println("Current url being checked: " + url);
                    try {
                        doc = fetchHTMLDocument(url);

                        emails = checkForEmails(doc, p);

                    } catch (IOException e) {
                        e.printStackTrace();
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

    private Document fetchHTMLDocument(String url) throws IOException {
        Connection.Response response = Jsoup.connect(url)
                .header("Accept-Encoding", "gzip, deflate")
                .userAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36")
                .referrer("http://www.google.com")
                .maxBodySize(0)
                .timeout(80000)
                .followRedirects(true)
                .execute();
        return response.parse();

    }

    private HashSet<String> checkForEmails(Document doc, Pattern p) {
        HashSet<String> emails = new HashSet<>();
        Matcher matcher;

        String stringMatcher = "empty";

        if (doc.head() != null) {
            stringMatcher = doc.head().toString();
        }
        if (doc.body() != null) {
            stringMatcher = stringMatcher + doc.body().toString();
        }

        if (doc.text() != null) {
            stringMatcher = stringMatcher + doc.text();
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
            int startingRow = scanner.nextInt();
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
