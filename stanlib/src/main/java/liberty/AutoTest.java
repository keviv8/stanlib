package liberty;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.*;

import javax.xml.bind.SchemaOutputResolver;
import java.io.*;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.InetSocketAddress;
import java.net.MalformedURLException;
import java.net.ProtocolException;
import java.net.Proxy;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.Format;
import java.text.ParseException;
import java.time.*;
import java.util.*;

import static liberty.ExcelUtil.*;
import static org.apache.poi.common.usermodel.HyperlinkType.URL;


public class AutoTest {
    public static Date realdate = new Date();
    public static LocalDate localDate = realdate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    public static int year = localDate.getYear();
    public static Month month = localDate.getMonth();
    public static int day = localDate.getDayOfMonth();
    public static boolean opfilecreated = false;
    public static String path;
    public static String daterange_path;
    public static String tempath;
    public static String jdrivecurrentyear = "\\\\libfin01\\libfin\\Libfin\\Systems\\UniCalc\\Production\\A2 Archive\\" + year + "\\";
    public static String jenkins_workspace = "C:\\Users\\vzk1008\\.jenkins\\workspace\\";
    public static String singleday_jenkins_projectname = "Daily-Run_Source-to-Test-Template_Single-Date\\";
    public static String daterange_jenkins_projectname = "Daily-Run_Source-to-Template_Date-Range\\";
    public static String inputFile = "inputDate(s).txt";
    public static String pathminusone;
    public static String tempathminusone;
    public static String date;
    public static String dateminusone;
    public static String start_path;
    public static String start_minusone_path;
    public static String end_minusone_path;
    public static String end_path;
    public static String product = null;
    public static String daterange_product = null;
    public static ArrayList < String > daysgone = new ArrayList < > ();
    public static String start_date;
    public static String start_date_minusone;
    public static String end_date;
    public static String SingleTransactionID = "";

    public static ArrayList < String > daterange = new ArrayList < > ();

    public static ArrayList < String > irs_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > irs_book_value = new ArrayList < > ();
    public static ArrayList < String > irs_pay_outstanding_notional = new ArrayList < > ();
    public static ArrayList < String > irs_receive_outstanding_notional = new ArrayList < > ();
    public static ArrayList < String > irs_accrued_income_value = new ArrayList < > ();
    public static ArrayList < String > irs_pay_accrued_income = new ArrayList < > ();
    public static ArrayList < String > irs_receive_accrued_income = new ArrayList < > ();
    public static ArrayList < String > irs_market_value = new ArrayList < > ();
    public static ArrayList < String > irs_market_valuet1 = new ArrayList < > ();
    public static ArrayList < String > irs_appreciation_value = new ArrayList < > ();
    public static ArrayList < String > irs_book_name = new ArrayList < > ();
    public static ArrayList < String > irs_realised_cash_flow = new ArrayList < > ();
    public static ArrayList < String > irs_unrealised_surplus_value = new ArrayList < > ();

    public static ArrayList < String > irn_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > irn_outstanding_notional = new ArrayList < > ();
    public static ArrayList < String > irn_market_value = new ArrayList < > ();
    public static ArrayList < String > irn_market_valuet1 = new ArrayList < > ();
    public static ArrayList < String > irn_accrued_income = new ArrayList < > ();
    public static ArrayList < String > irn_realised_cash_flow = new ArrayList < > ();
    public static ArrayList < String > irn_book_value = new ArrayList < > ();
    public static ArrayList < String > irn_appreciation_value = new ArrayList < > ();
    public static ArrayList < String > irn_income_value = new ArrayList < > ();
    public static ArrayList < String > irn_bookname = new ArrayList < > ();
    public static ArrayList < String > irn_unrealised_surplus_value = new ArrayList < > ();

    public static ArrayList < String > crn_book_value = new ArrayList < > ();
    public static ArrayList < String > crn_market_value = new ArrayList < > ();
    public static ArrayList < String > crn_accrued_income = new ArrayList < > ();
    public static ArrayList < String > crn_appreciation_value = new ArrayList < > ();
    public static ArrayList < String > crn_unreal_surplus = new ArrayList < > ();
    public static ArrayList < String > crn_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > crn_book_name = new ArrayList < > ();
    public static ArrayList < String > crn_market_valuet1 = new ArrayList < > ();
    public static ArrayList < String > crn_realised_cash_flow = new ArrayList < > ();

    public static ArrayList < String > eln_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > eln_book_name = new ArrayList < > ();
    public static ArrayList < String > eln_book_value = new ArrayList < > ();
    public static ArrayList < String > eln_appreciation_value = new ArrayList < > ();
    public static ArrayList < String > eln_unreal_surplus = new ArrayList < > ();
    public static ArrayList < String > eln_market_value = new ArrayList < > ();
    public static ArrayList < String > eln_market_valuet1 = new ArrayList < > ();


    public static ArrayList < String > cash_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > cash_transaction_reference_new = new ArrayList < > ();
    public static ArrayList < String > cash_book_name = new ArrayList < > ();
    public static ArrayList < String > cash_book_name_new = new ArrayList < > ();
    public static ArrayList < String > cash_book_value = new ArrayList < > ();
    public static ArrayList < String > cash_accrued_income = new ArrayList < > ();
    public static ArrayList < String > cash_income = new ArrayList < > ();

    public static ArrayList < String > eqf_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > eqf_book_name = new ArrayList < > ();
    public static ArrayList < String > eqf_outstanding_income = new ArrayList < > ();
    public static ArrayList < String > eqf_income = new ArrayList < > ();

    public static ArrayList < String > eqfwd_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > eqfwd_book_name = new ArrayList < > ();
    public static ArrayList < String > eqfwd_appreciation_value = new ArrayList < > ();
    public static ArrayList < String > eqfwd_appreciation_valuet1 = new ArrayList < > ();
    public static ArrayList < String > eqfwd_income_value = new ArrayList < > ();
    public static ArrayList < String > eqfwd_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > cfs_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > cfs_book_name = new ArrayList < > ();
    public static ArrayList < String > cfs_appreciation_value = new ArrayList < > ();
    public static ArrayList < String > cfs_appreciation_valuet1 = new ArrayList < > ();
    public static ArrayList < String > cfs_income_value = new ArrayList < > ();
    public static ArrayList < String > cfs_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > irfwd_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > irfwd_book_name = new ArrayList < > ();
    public static ArrayList < String > irfwd_appreciation_value = new ArrayList < > ();
    public static ArrayList < String > irfwd_appreciation_valuet1 = new ArrayList < > ();
    public static ArrayList < String > irfwd_income_value = new ArrayList < > ();
    public static ArrayList < String > irfwd_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > instruments = new ArrayList < > ();

    ////declarations for data range columns for 9 Instruments
    public static ArrayList < String > irn_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > irn_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > irn_daterange_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > crn_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > crn_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > crn_daterange_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > irs_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > irs_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > irs_daterange_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > eln_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > eln_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > eln_daterange_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > eqfwd_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > eqfwd_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > eqfwd_daterange_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > cfs_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > cfs_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > cfs_daterange_unrealised_surplus = new ArrayList < > ();

    public static ArrayList < String > eqf_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > eqf_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > eqf_daterange_unrealised_surplus = new ArrayList < > ();
    public static ArrayList < String > cash_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > cash_daterange_transaction_reference_new = new ArrayList < > ();
    public static ArrayList < String > cash_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > cash_daterange_unrealised_surplus = new ArrayList < > ();
    public static ArrayList < String > irfwd_daterange_transaction_reference = new ArrayList < > ();
    public static ArrayList < String > irfwd_daterange_income_value = new ArrayList < > ();
    public static ArrayList < String > irfwd_daterange_unrealised_surplus = new ArrayList < > ();



    public static void main(String[] args) throws Exception {



        //temp declarations - to be removed later
        //        System.out.println("Would you like to enter\n1] A date\n2] A date range");
        //        Scanner s = new Scanner(System.in);
        //        String choice = s.nextLine();
        //
        //        System.out.println("Enter the date in yyyymmdd format");
        //        String date = s.nextLine();
        //        String filename = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report 20190401.xlsx";
        Instant start = Instant.now();
        //        if(choice.equals("1"))
        //        path = path + date;
        //        IRSFlow();

        LocalDate that = LocalDate.now();
        System.out.println("today is " + LocalDate.now().minusDays(1) + " and comparing it with " + that.toString());
        for (int i = 0; i < 1000; i++) {
            daysgone.add(LocalDate.now().minusDays(i).toString().replace("-", ""));
        }
        System.out.println(daysgone);
        System.out.println(getDateRanges("20190531", "20190404"));


        BigDecimal one = new BigDecimal(180963492.739);
        BigDecimal two = new BigDecimal(-180918250.128);
        String bongu = "";
        String thongo = "sdgdg";
        int minussss = -1;
        System.out.println(one.subtract(two));
        System.out.println((one.subtract(two).multiply(new BigDecimal(minussss))));
        //        System.out.println("Is BigDecimal? - " + isNumeric(String.valueOf(one)));
        //        System.out.println("Is BigDecimal? - " + isNumeric(String.valueOf(two)));
        //        System.out.println("Is BigDecimal? - " + isNumeric(bongu));
        //        System.out.println("Is BigDecimal? - " + isNumeric(thongo));
        readInputFile();
        System.out.println("Valid Business Days in between date range are " + getBusinessDaysDateRanges());


        //        System.out.println(months);


        //        csvToXLSX(path,"Static Contract Data Report 20190329.csv");
        //        System.out.println("CSV conversion done, waiting for a couple seconds");
        //        ArrayList<String> files = getFilenamesFromFolder(path);
        //        System.out.println(files);
        //        Thread.sleep(9000);

        //        readSourceFile(path, filename, sheet);

    }

    @BeforeMethod
    public static void beforeMethodSingleDay() {
        System.out.println("beforeMethodSingleDay");
        path = "C:\\Users\\vzk1008\\Documents\\04 production\\";
        date = "20190531";
        pathminusone = path;
        dateminusone = "20190530";
        path = path + date + "\\";
        tempath = path;
        pathminusone = pathminusone + dateminusone;
        System.out.println("Path " + path);
    }

    @BeforeMethod
    public static void beforeMethodDateRange() {
        daterange_path = "C:\\Users\\vzk1008\\Documents\\04 production\\";
        start_path = daterange_path + start_date + "\\";
        end_path = daterange_path + end_date + "\\";
        start_minusone_path = daterange_path + start_date_minusone + "\\";
        System.out.println("Start path = " + start_path + "\nEnd Path = " + end_path);
    }


    public static float timediff(Instant from, Instant to) {
        return ((float)(Duration.between(from, to).toMillis()) / (float) 1000);
    }

    public static String today() {
        return daysgone.get(0);
    }

    public static String yesterday() {
        return daysgone.get(1);
    }



    public static boolean isNumeric(String strNum) {
        try {
            double d = Double.parseDouble(strNum);
        } catch (NumberFormatException | NullPointerException nfe) {
            return false;
        }
        return true;
    }

    public static List < String > getDateRanges(String from, String to) {
        ArrayList < String > dr = new ArrayList < > ();
        int from_index, to_index;
        for (String s: daysgone) {
            if (s.equalsIgnoreCase(from))
                from_index = daysgone.indexOf(from);
        }

        return daysgone.subList(daysgone.indexOf(from), daysgone.indexOf(to) + 1);
    }


    public static String readInputFile() throws FileNotFoundException {
        System.out.println("Reading input file");
        String line = null;
        for (Scanner sc = new Scanner(new File(jenkins_workspace + daterange_jenkins_projectname + inputFile)); sc.hasNext();) {
            line = sc.nextLine();
        }
        line.replaceAll("\n", "");
        System.out.println(line);
        if (line.contains(",")) {
            start_date = line.substring(0, 8);
            end_date = line.substring(line.indexOf(",") + 1);
            System.out.println("Start Date = " + start_date + "\nEnd Date = " + end_date);
        } else {
            start_date = line;
            end_date = null;
        }
        return line;
    }


    public static String thatdaysyesterdayDate(String thatday) throws FileNotFoundException {
        readInputFile();
        ArrayList < String > businessdaysrange = new ArrayList < > ();
        List < String > exactbusinessdays = null;
        ArrayList < String > eachmonth = new ArrayList < > ();
        System.out.println(start_date + "\t" + end_date);
        try {
            //        System.out.println("Reading from Jdrive location - "+jdrivecurrentyear);
            //        System.out.println("Folder names in Jdrive current are "+getFolderNamesFromDirectory(jdrivecurrentyear));
            int i = 0;
            for (String s: getFolderNamesFromDirectory(jdrivecurrentyear)) {
                eachmonth = getFolderNamesFromDirectory(s + "\\");
                //            System.out.println("Sub folders in each month in 2019 Jdrive are "+getFolderNamesFromDirectory(s+"\\"));
                for (int j = 0; j < eachmonth.size(); j++)
                    businessdaysrange.add(eachmonth.get(j).substring(eachmonth.get(j).length() - 8));
            }
            //        Collections.sort(businessdaysrange,Collections.reverseOrder());
            Collections.sort(businessdaysrange);
            //        System.out.println("Valid Business Days range retrieved from JDrive are  "+businessdaysrange);
            exactbusinessdays = businessdaysrange.subList(businessdaysrange.indexOf(start_date), businessdaysrange.indexOf(end_date) + 1);
            start_date_minusone = businessdaysrange.get(businessdaysrange.indexOf(start_date) - 1);
            System.out.println(thatday + "'s yesterday is " + businessdaysrange.get(businessdaysrange.indexOf(thatday) - 1));
            System.out.println(daterange_path + "\\" + businessdaysrange.get(businessdaysrange.indexOf(thatday) - 1));
        } catch (IndexOutOfBoundsException e) {
            System.out.println("ENTERED DATES ARE NOT BUSINESS DAYS");
            e.printStackTrace();
        }
        return businessdaysrange.get(businessdaysrange.indexOf(thatday) - 1);
    }


    public static String thatdaysyesterday(String thatday) throws FileNotFoundException {
        readInputFile();
        ArrayList < String > businessdaysrange = new ArrayList < > ();
        List < String > exactbusinessdays = null;
        ArrayList < String > eachmonth = new ArrayList < > ();
        System.out.println(start_date + "\t" + end_date);
        try {
            //        System.out.println("Reading from Jdrive location - "+jdrivecurrentyear);
            //        System.out.println("Folder names in Jdrive current are "+getFolderNamesFromDirectory(jdrivecurrentyear));
            int i = 0;
            for (String s: getFolderNamesFromDirectory(jdrivecurrentyear)) {
                eachmonth = getFolderNamesFromDirectory(s + "\\");
                //            System.out.println("Sub folders in each month in 2019 Jdrive are "+getFolderNamesFromDirectory(s+"\\"));
                for (int j = 0; j < eachmonth.size(); j++)
                    businessdaysrange.add(eachmonth.get(j).substring(eachmonth.get(j).length() - 8));
            }
            //        Collections.sort(businessdaysrange,Collections.reverseOrder());
            Collections.sort(businessdaysrange);
            //        System.out.println("Valid Business Days range retrieved from JDrive are  "+businessdaysrange);
            exactbusinessdays = businessdaysrange.subList(businessdaysrange.indexOf(start_date), businessdaysrange.indexOf(end_date) + 1);
            start_date_minusone = businessdaysrange.get(businessdaysrange.indexOf(start_date) - 1);
            System.out.println(thatday + "'s yesterday is " + businessdaysrange.get(businessdaysrange.indexOf(thatday) - 1));
            System.out.println(daterange_path + "\\" + businessdaysrange.get(businessdaysrange.indexOf(thatday) - 1));
        } catch (IndexOutOfBoundsException e) {
            System.out.println("ENTERED DATES ARE NOT BUSINESS DAYS");
            e.printStackTrace();
        }
        return daterange_path + "\\" + businessdaysrange.get(businessdaysrange.indexOf(thatday) - 1);
    }


    @BeforeSuite
    public static List < String > getBusinessDaysDateRanges() throws FileNotFoundException {
        readInputFile();
        ArrayList < String > businessdaysrange = new ArrayList < > ();
        List < String > exactbusinessdays = null;
        ArrayList < String > eachmonth = new ArrayList < > ();
        System.out.println(start_date + "\t" + end_date);
        try {
                    System.out.println("Reading from Jdrive location - "+jdrivecurrentyear);
                    System.out.println("Folder names in Jdrive current are "+getFolderNamesFromDirectory(jdrivecurrentyear));
            int i = 0;
            for (String s: getFolderNamesFromDirectory(jdrivecurrentyear)) {
                eachmonth = getFolderNamesFromDirectory(s + "\\");
                            System.out.println("Sub folders in each month in 2019 Jdrive are "+getFolderNamesFromDirectory(s+"\\"));
                for (int j = 0; j < eachmonth.size(); j++)
                    businessdaysrange.add(eachmonth.get(j).substring(eachmonth.get(j).length() - 8));
            }
//                    Collections.sort(businessdaysrange,Collections.reverseOrder());
            Collections.sort(businessdaysrange);
                    System.out.println("Valid Business Days range retrieved from JDrive are  "+businessdaysrange);
            exactbusinessdays = businessdaysrange.subList(businessdaysrange.indexOf(start_date), businessdaysrange.indexOf(end_date) + 1);
            start_date_minusone = businessdaysrange.get(businessdaysrange.indexOf(start_date) - 1);
        } catch (IndexOutOfBoundsException e) {
            System.out.println("ENTERED DATES ARE NOT BUSINESS DAYS");
            e.printStackTrace();
        }
        for (String s: exactbusinessdays) {
            daterange.add(s);
        }
        System.out.println("Final Date Range for which testing to be done " + daterange);
        return exactbusinessdays;
    }


    @Test
    public static void fetchCashMIforDateRange() throws Exception {

        Instant start = Instant.now();
        daterange_product = "CashMI";
        instruments.add(daterange_product);
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        String filename1 = "Cash_MI_Daily_PLA-2019" + end_date + ".csv";
        for (String s: files) {
            if (s.contains("Cash_MI_Daily_PLA-2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Cash_MI_Daily_PLA"));
        }
        ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String identifier = "CashMI";

        csvToXLSX(end_path, filename1);

        ArrayList < String > comparator2, comparator3, comparator4, comparator5;

        comparator2 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 19)); // book value
        comparator3 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 20)); // accrued income
        comparator4 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 4)); // book name
        comparator5 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 3)); // primary key

        for (int i = 2; i < comparator5.size(); i++) {
            cash_daterange_transaction_reference.add(comparator5.get(i));
            cash_book_name.add(comparator4.get(i));
            cash_book_value.add(comparator2.get(i));
            cash_accrued_income.add(comparator3.get(i));
            //            cash_income.add(comparator1.get(i));
        }



        //////////////////////////////////////CODE FOR DOING INCOME VALUE
        ArrayList < String > first_income  ;
        ArrayList < String > second_income ;
        ArrayList < String > sum_income;
        ArrayList < String > total_income = new ArrayList < > ();
        String firstpath ;
        String secondpath ;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_income = getCashMIIncomeValueforDateRange(firstpath);
                second_income = getCashMIIncomeValueforDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);
            }
            System.out.println("Final Total Income for even size array is " + total_income.size() + " and the array is ");
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_income = getCashMIIncomeValueforDateRange(firstpath);
                second_income = getCashMIIncomeValueforDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_income = getCashMIIncomeValueforDateRange(firstpath);
                    total_income = addThoseArrays(total_income, first_income);
                }
            }
            //            System.out.println("Transaction count size is " + eqf_daterange_transaction_reference.size() + " and the list is " + eqf_daterange_transaction_reference);
            System.out.println("Final Total Income for uneven size array is " + total_income.size() + " and the array is ");
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING INCOME
        cash_daterange_income_value = total_income;
        cash_daterange_income_value = negateThatArray(cash_daterange_income_value);


        for (String s: cash_daterange_transaction_reference) {
            String n = s.replace("\"", "");
            cash_daterange_transaction_reference_new.add(n);
        }
//        System.out.println("Cash Date Range Old " + cash_daterange_transaction_reference);
//        System.out.println("Cash Date Range New " + cash_daterange_transaction_reference_new);

        for (String s: cash_book_name) {
            String n = s.replace("\"", "");
            cash_book_name_new.add(n);
        }

        System.out.println("Date range for " + daterange_product + " Transaction Reference size is " + cash_daterange_transaction_reference_new.size() + " and the array is ");
        //        System.out.println("Book Name size is " + cash_book_name.size() + " and the array is " + cash_book_name);
        System.out.println("Date range for " + daterange_product + " Accrued Income size is " + cash_accrued_income.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Book Value size is " + cash_book_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Income size is " + cash_daterange_income_value.size() + " and the array is ");

        Instant end = Instant.now();
        System.out.println("Time taken for CashMI flow - " + timediff(start, end) + " seconds");

    }

    @Test
    public static void fetchEQFUTforDateRange() throws Exception {
        Instant start = Instant.now();
        daterange_product = "EQFUT";
        instruments.add(daterange_product);
        String filename = "Static Contract Data Report.csv";
        String filename2 = "AB - PROD - FIN - BO Libfin 10D Expected CashFlows Report" + end_date + ".csv";
        String f2sheetname = "AB - PROD - FIN - BO Libfin 10D";
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        for (String s: files) {
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
            if (s.contains("AB - PROD - FIN - BO Libfin 10D Expected CashFlows Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        }
        String fn_sheetname = "Static Contract Data Report 201";

        ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + thatdaysyesterdayDate(end_date) + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        String identifier = "EQFUT";

        //        System.out.println(end_path+"\t"+filename);
        csvToXLSX(end_path, filename);
        //        System.out.println(end_path+"\t"+filename1);
        csvToXLSX(end_path, filename2);
        //        System.out.println(end_path+"\t"+filename3);
        csvToXLSX(end_minusone_path, filename4);

        System.out.println("reading primary key from file " + filename2 + "at path " + end_path + " and sheetname " + f2sheetname + " with identifier " + identifier);
        eqf_daterange_transaction_reference = readPrimaryKey(end_path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_abprod, comparator1;
        //        System.out.println("Transaction count size is " + eqf_daterange_transaction_reference.size() + " and the list is " + eqf_daterange_transaction_reference);
        comparator1 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 7)); //outstanding income
        fkey_abprod = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 3));
        //        System.out.println("Comparat   or 1 size is " + comparator1.size() + " and the list is " + comparator1);

        ////// POPULATION OF OUTSTANDING INCOME
        for (int i = 0; i < eqf_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (eqf_daterange_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    eqf_outstanding_income.add(comparator1.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    eqf_outstanding_income.add("0.00");
                }
            }
        }


        //////////////////////////////////////CODE FOR DOING INCOME VALUE
        ArrayList < String > first_income ;
        ArrayList < String > second_income;
        ArrayList < String > sum_income ;
        ArrayList < String > total_income = new ArrayList < > ();
        String firstpath = null;
        String secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_income = getEQFUTIncomeValueForDateRange(firstpath);
                second_income = getEQFUTIncomeValueForDateRange(secondpath);
//                                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
//                                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);
                prettyPrint(eqf_daterange_transaction_reference,first_income,second_income,sum_income);
            }
            //            System.out.println("Final Total Income for even size array is "+total_income.size()+" and the array is "+total_income);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_income = getEQFUTIncomeValueForDateRange(firstpath);
                second_income = getEQFUTIncomeValueForDateRange(secondpath);
//                                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
//                                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_income = getEQFUTIncomeValueForDateRange(firstpath);
                    total_income = addThoseArrays(total_income, first_income);
                }
                prettyPrint(eqf_daterange_transaction_reference,first_income,second_income,sum_income);
            }
            //            System.out.println("Transaction count size is " + eqf_daterange_transaction_reference.size() + " and the list is " + eqf_daterange_transaction_reference);
            System.out.println("Final Total Income for uneven size array is " + total_income.size() + " and the array is " + total_income);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING INCOME
        eqf_daterange_income_value = negateThatArray(total_income);
        System.out.println("Date range for " + daterange_product + "  Transaction reference  size is " + eqf_daterange_transaction_reference.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Book Name sizeis " + eqf_book_name.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + "  Outstanding income size is " + eqf_outstanding_income.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + "  Income size is " + eqf_daterange_income_value.size() + " and the array is ");
        //        createOneBigDateRangeOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for EQF flow - " + timediff(start, end) + " seconds");

    }

    @Test
    public static void fetchEQFWDforDateRange() throws Exception {

        Instant start = Instant.now();
        daterange_product = "EQFWD";
        instruments.add(daterange_product);
        String filename = "Static Contract Data Report.csv";
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        //        for (String s: files)
        //            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
        //                filename = s.substring(s.indexOf("Static"));
        //                filename = "Static Contract Data Report "+date+".csv";
        //        String filename1 = "Data_Trade_ELN(A)_Stat_" + end_date + ".csv";
        //        for (String s: files)
        //            if (s.contains("Data_Trade_ELN(A)_Stat_") && s.contains(".csv"))
        //                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + end_date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + end_date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "EQFWD";

        ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + thatdaysyesterdayDate(end_date) + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //        System.out.println(end_path+"\t"+filename);
        //        csvToXLSX(end_path, filename);
        //        System.out.println(end_path+"\t"+filename1);
        //        csvToXLSX(end_path, filename1);
        //        System.out.println(end_path+"\t"+filename2);
        csvToXLSX(end_path, filename2);
        //        System.out.println(end_path+"\t"+filename3);
        csvToXLSX(end_path, filename3);
        //        System.out.println(start_path+"\t"+filename4);
        csvToXLSX(end_minusone_path, filename4);


        ArrayList < String > fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6;

        comparator1 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //primary key
        comparator2 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 1)); //bookname
        comparator3 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //appreciation value//i column in ab prod
        comparator4 = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //income value//multiply this with -1
        comparator5 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's appreciation value
        fk_realizedcash = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //foreign key in realised cashflows file
        fkey_abprod = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //fk in ab prod file//keep this as backup
        fkey_abprod_yesterday = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 5)); //
        comparator6 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //

        for (int i = 0; i < comparator1.size(); i++) { //filter primary keys
            if (comparator1.get(i).contains("EQFWD")) {
                eqfwd_daterange_transaction_reference.add(comparator1.get(i));
            }
        }


        ////// POPULATION OF BOOK VALUE and APPRECIATION VALUE
        for (int i = 0; i < eqfwd_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < comparator1.size(); j++) {
                if (eqfwd_daterange_transaction_reference.get(i).equals(comparator1.get(j))) {
                    eqfwd_book_name.add(comparator2.get(j));
                    eqfwd_appreciation_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == comparator1.size() - 1) {
                    eqfwd_book_name.add("0.00");
                    eqfwd_appreciation_value.add("0.00");
                }
            }
        }

        //YESTERDAY'S APPRECIATION VALUE
        for (int i = 0; i < eqfwd_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (eqfwd_daterange_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    eqfwd_appreciation_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    eqfwd_appreciation_valuet1.add("0.00");
                }
            }
        }

        //        ////// POPULATION OF INCOME VALUE
        //        for (int i = 0; i < eqfwd_daterange_transaction_reference.size(); i++) {
        //            boolean match_found = false;
        //            for (int j = 0; j < fk_realizedcash.size(); j++) {
        //                if (eqfwd_daterange_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
        //                    eqfwd_income_value.add(comparator4.get(j));
        //                    match_found = true;
        //                }
        //                if (!match_found && j == fk_realizedcash.size() - 1) {
        //                    eqfwd_income_value.add("0.00");
        //                }
        //            }
        //        }

        //        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        //        BigDecimal appvalue_bd, appvaluet1_bd, unrealised_result;
        //        for (int i = 0; i < eqfwd_daterange_transaction_reference.size(); i++) {
        //            try {
        //                appvalue_bd = new BigDecimal(eqfwd_appreciation_value.get(i));
        //                appvaluet1_bd = new BigDecimal(eqfwd_appreciation_valuet1.get(i));
        //                unrealised_result = appvalue_bd.subtract(appvaluet1_bd);
        //                eqfwd_unrealised_surplus.add(String.valueOf(unrealised_result)); ////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
        ////                    eqfwd_unrealised_surplus.add(eqfwd_appreciation_value.get(i));
        //            } catch (Exception e) {
        //                //                e.printStackTrace();
        //                eqfwd_unrealised_surplus.add("ERROR");
        //            }
        //        }

        //////////////////////////////////////CODE FOR DOING UNREAL VALUE
        ArrayList < String > first_unreal = new ArrayList < > ();
        ArrayList < String > second_unreal = new ArrayList < > ();
        ArrayList < String > sum_unreal = new ArrayList < > ();
        ArrayList < String > total_unreal = new ArrayList < > ();
        String firstpath = null, secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                //                System.out.println("Sum unreal size is "+sum_unreal.size()+" and the array is "+sum_unreal);
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);
            }
            //            System.out.println("Final Total unreal for even size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is " + sum_unreal);
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                    System.out.println("First unreal size is " + first_unreal.size() + " and the array is " + first_unreal);
                    total_unreal = addThoseArrays(total_unreal, first_unreal);
                }
            }
            //            System.out.println("Transaction count size is " + eqfwd_daterange_transaction_reference.size() + " and the list is " + eqfwd_daterange_transaction_reference);
            //            System.out.println("Final Total unreal for uneven size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING unreal
        eqfwd_daterange_unrealised_surplus = total_unreal;


        //////////////////////////////////////CODE FOR DOING INCOME VALUE
        ArrayList < String > first_income = new ArrayList < > ();
        ArrayList < String > second_income = new ArrayList < > ();
        ArrayList < String > sum_income = new ArrayList < > ();
        ArrayList < String > total_income = new ArrayList < > ();
        firstpath = null;
        secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                System.out.println("First income size is " + first_income.size() + " and the array is ");
                System.out.println("Second income size is " + second_income.size() + " and the array is ");
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);
            }
            //            System.out.println("Final Total Income for even size array is "+total_income.size()+" and the array is "+total_income);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_income = getIncomeValueForDateRange(firstpath);
                    total_income = addThoseArrays(total_income, first_income);
                }
            }
            //            System.out.println("Transaction count size is " + eqfwd_daterange_transaction_reference.size() + " and the list is " + eqfwd_daterange_transaction_reference);
            //            System.out.println("Final Total Income for uneven size array is "+total_income.size()+" and the array is "+total_income);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING INCOME

        eqfwd_daterange_income_value = total_income;



        eqfwd_daterange_income_value = negateThatArray(eqfwd_daterange_income_value);
        eqfwd_daterange_unrealised_surplus = negateThatArray(eqfwd_daterange_unrealised_surplus);


        System.out.println("Date range for " + daterange_product + "  Transaction count size is " + eqfwd_daterange_transaction_reference.size() + " and the list is ");
        //        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        //        System.out.println("Book Value size is " + crn_book_value.size() + " and the array is " + crn_book_value);
        //        System.out.println("Accrued Income Value is " + crn_accrued_income.size() + " and the array is " + crn_accrued_income);
        System.out.println("Date range for " + daterange_product + "  Book Name size is " + eqfwd_book_name.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + "  Appreciation Value size is " + eqfwd_appreciation_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + "  Income Value " + eqfwd_daterange_income_value.size() + " and the array is ");
        //        System.out.println("Realised Cash Flow or Income Value size is " + crn_realised_cash_flow.size() + " and the array is " + crn_realised_cash_flow);
        System.out.println("Date range for " + daterange_product + "  Unrealised Surplus Value size is " + eqfwd_daterange_unrealised_surplus.size() + " and the array is ");

        //        createOneBigDateRangeOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for EQFWD flow - " + timediff(start, end) + " seconds");


    }

    @Test
    public static void fetchCFSforDateRange() throws Exception {

        Instant start = Instant.now();
        daterange_product = "CFS";
        instruments.add(daterange_product);
        String filename = "Static Contract Data Report.csv";
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        //        filename = "Static Contract Data Report "+date+".csv";
        //        String filename1 = "Data_Trade_ELN(A)_Stat_" + end_date + ".csv";
        //        for (String s: files)
        //            if (s.contains("Data_Trade_ELN(A)_Stat_") && s.contains(".csv"))
        //                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + end_date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + end_date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "ELN(A)";

        ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + thatdaysyesterdayDate(end_date) + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //        System.out.println(end_path+"\t"+filename);
        //        csvToXLSX(end_path, filename);
        //        System.out.println(end_path+"\t"+filename1);
        //        csvToXLSX(end_path, filename1);
        //        System.out.println(end_path+"\t"+filename2);
        csvToXLSX(end_path, filename2);
        //        System.out.println(end_path+"\t"+filename3);
        csvToXLSX(end_path, filename3);
        //        System.out.println(start_path+"\t"+filename4);
        csvToXLSX(end_minusone_path, filename4);

        ArrayList < String > fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6;

        comparator1 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //primary key
        comparator2 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 1)); //bookname
        comparator3 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //appreciation value//i column in ab prod
        comparator4 = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //income value//multiply this with -1
        comparator5 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's appreciation value
        fk_realizedcash = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //foreign key in realised cashflows file
        fkey_abprod = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //fk in ab prod file//keep this as backup
        fkey_abprod_yesterday = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 5)); //
        comparator6 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //

        for (int i = 0; i < comparator1.size(); i++) { //filter primary keys
            if (comparator1.get(i).contains("CFS")) {
                cfs_daterange_transaction_reference.add(comparator1.get(i));
            }
        }


        ////// POPULATION OF BOOK VALUE and APPRECIATION VALUE
        for (int i = 0; i < cfs_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < comparator1.size(); j++) {
                if (cfs_daterange_transaction_reference.get(i).equals(comparator1.get(j))) {
                    cfs_book_name.add(comparator2.get(j));
                    cfs_appreciation_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == comparator1.size() - 1) {
                    cfs_book_name.add("0.00");
                    cfs_appreciation_value.add("0.00");
                }
            }
        }

        //YESTERDAY'S APPRECIATION VALUE
        for (int i = 0; i < cfs_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (cfs_daterange_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    cfs_appreciation_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    cfs_appreciation_valuet1.add("0.00");
                }
            }
        }

        //        ////// POPULATION OF INCOME VALUE
        //        for (int i = 0; i < cfs_transaction_reference.size(); i++) {
        //            boolean match_found = false;
        //            for (int j = 0; j < fk_realizedcash.size(); j++) {
        //                if (cfs_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
        //                    cfs_income_value.add(comparator4.get(j));
        //                    match_found = true;
        //                }
        //                if (!match_found && j == fk_realizedcash.size() - 1) {
        //                    cfs_income_value.add("0.00");
        //                }
        //            }
        //        }

        //        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        //        BigDecimal appvalue_bd, appvaluet1_bd, unrealised_result;
        //        for (int i = 0; i < cfs_transaction_reference.size(); i++) {
        //            try {
        //                appvalue_bd = new BigDecimal(cfs_appreciation_value.get(i));
        //                appvaluet1_bd = new BigDecimal(cfs_appreciation_valuet1.get(i));
        //                unrealised_result = appvalue_bd.subtract(appvaluet1_bd);
        //                cfs_unrealised_surplus.add(String.valueOf(unrealised_result));////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
        ////                cfs_unrealised_surplus.add(cfs_appreciation_value.get(i));
        //            } catch (Exception e) {
        //                //                e.printStackTrace();
        //                cfs_unrealised_surplus.add("ERROR");
        //            }
        //        }



        //////////////////////////////////////CODE FOR DOING UNREAL VALUE
        ArrayList < String > first_unreal = new ArrayList < > ();
        ArrayList < String > second_unreal = new ArrayList < > ();
        ArrayList < String > sum_unreal = new ArrayList < > ();
        ArrayList < String > total_unreal = new ArrayList < > ();
        String firstpath = null, secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is ");
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);
            }
            //            System.out.println("Final Total unreal for even size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is " );
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                    System.out.println("First unreal size is " + first_unreal.size() + " and the array is " );
                    total_unreal = addThoseArrays(total_unreal, first_unreal);
                }
            }
            //            System.out.println("Transaction count size is " + cfs_daterange_transaction_reference.size() + " and the list is " + cfs_daterange_transaction_reference);
            //            System.out.println("Final Total unreal for uneven size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING unreal
        cfs_daterange_unrealised_surplus = total_unreal;


        //////////////////////////////////////CODE FOR DOING INCOME VALUE
        ArrayList < String > first_income = new ArrayList < > ();
        ArrayList < String > second_income = new ArrayList < > ();
        ArrayList < String > sum_income = new ArrayList < > ();
        ArrayList < String > total_income = new ArrayList < > ();
        firstpath = null;
        secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);
            }
            //            System.out.println("Final Total Income for even size array is "+total_income.size()+" and the array is "+total_income);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_income = getIncomeValueForDateRange(firstpath);
                    total_income = addThoseArrays(total_income, first_income);
                }
            }
            //            System.out.println("Transaction count size is " + cfs_daterange_transaction_reference.size() + " and the list is " + cfs_daterange_transaction_reference);
            //            System.out.println("Final Total Income for uneven size array is "+total_income.size()+" and the array is "+total_income);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING INCOME

        cfs_daterange_income_value = total_income;

        cfs_daterange_income_value = negateThatArray(cfs_daterange_income_value);
        cfs_daterange_unrealised_surplus = negateThatArray(cfs_daterange_unrealised_surplus);

        System.out.println("Date range for " + daterange_product + " Transaction count size is " + cfs_daterange_transaction_reference.size() + " and the list is ");
        //        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        //        System.out.println("Book Value size is " + crn_book_value.size() + " and the array is " + crn_book_value);
        //        System.out.println("Accrued Income Value is " + crn_accrued_income.size() + " and the array is " + crn_accrued_income);
        System.out.println("Date range for " + daterange_product + "  Book Name size is " + cfs_book_name.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Appreciation Value size is " + cfs_appreciation_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + "  Income Value " + cfs_daterange_income_value.size() + " and the array is ");
        //        System.out.println("Realised Cash Flow or Income Value size is " + crn_realised_cash_flow.size() + " and the array is " + crn_realised_cash_flow);
        System.out.println("Date range for " + daterange_product + "  Unrealised Surplus Value size is " + cfs_daterange_unrealised_surplus.size() + " and the array is ");

        //        createOneBigDateRangeOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for CFS flow - " + timediff(start, end) + " seconds");


    }

    @Test
    public static void fetchELNforDateRange() throws Exception {
        Instant start = Instant.now();
        daterange_product = "ELN(A)";
        instruments.add(daterange_product);
        String filename = "Static Contract Data Report.csv";
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        //        filename = "Static Contract Data Report "+date+".csv";
        String filename1 = "Data_Trade_ELN(A)_Stat_" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_ELN(A)_Stat_") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + end_date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + end_date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "ELN(A)";

        ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + thatdaysyesterdayDate(end_date) + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //        System.out.println(end_path+"\t"+filename);
        csvToXLSX(end_path, filename);
        //        System.out.println(end_path+"\t"+filename1);
        csvToXLSX(end_path, filename1);
        //        System.out.println(end_path+"\t"+filename2);
        csvToXLSX(end_path, filename2);
        //        System.out.println(end_path+"\t"+filename3);
        csvToXLSX(end_path, filename3);
        //        System.out.println(start_path+"\t"+filename4);
        csvToXLSX(end_minusone_path, filename4);

        eln_daterange_transaction_reference = readPrimaryKey(end_path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5;
        fkey_datatrade = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 7)); //outstanding notional or bookvalue
        //        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 17)); // accrued intetest native
        comparator3 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column//i column in ab prod
        comparator4 = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //
        comparator5 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        //        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //
        fkey_abprod = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));


        ////// POPULATION OF BOOK VALUE
        for (int i = 0; i < eln_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (eln_daterange_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    eln_book_value.add(comparator1.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    eln_book_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < eln_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (eln_daterange_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    eln_market_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    eln_market_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < eln_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (eln_daterange_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    eln_market_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    eln_market_valuet1.add("0.00");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        BigDecimal marval_bd, unrealised_result, markett1_bd, bookval_bd, appvalue_result;
        for (int i = 0; i < eln_daterange_transaction_reference.size(); i++) {
            try {
                marval_bd = new BigDecimal(eln_market_value.get(i));
                markett1_bd = new BigDecimal(eln_market_valuet1.get(i));
                bookval_bd = new BigDecimal(eln_book_value.get(i));
                unrealised_result = marval_bd.subtract(markett1_bd);
                appvalue_result = marval_bd.subtract(bookval_bd);
                eln_appreciation_value.add(String.valueOf(appvalue_result));
                //                eln_unreal_surplus.add(String.valueOf(unrealised_result));////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                //                eln_unreal_surplus.add(eln_market_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();

                eln_appreciation_value.add("ERROR");
                //                eln_unreal_surplus.add("ERROR");
            }
        }


        //////////////////////////////////////CODE FOR DOING UNREAL VALUE
        ArrayList < String > first_unreal = new ArrayList < > ();
        ArrayList < String > second_unreal = new ArrayList < > ();
        ArrayList < String > sum_unreal = new ArrayList < > ();
        ArrayList < String > total_unreal = new ArrayList < > ();
        String firstpath = null, secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is " );
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is " );
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                //                System.out.println("Sum unreal size is "+sum_unreal.size()+" and the array is "+sum_unreal);
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);
            }
            //            System.out.println("Final Total unreal for even size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is ");
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                    System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                    total_unreal = addThoseArrays(total_unreal, first_unreal);
                }
            }
            //            System.out.println("Transaction count size is " + eln_daterange_transaction_reference.size() + " and the list is " + eln_daterange_transaction_reference);
            //            System.out.println("Final Total unreal for uneven size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING unreal
        eln_daterange_unrealised_surplus = total_unreal;

                //////////////////////////////////////CODE FOR DOING INCOME VALUE
                ArrayList<String> first_income = new ArrayList<>();
                ArrayList<String> second_income = new ArrayList<>();
                ArrayList<String> sum_income = new ArrayList<>();
                ArrayList<String> total_income = new ArrayList<>();
                firstpath=null;
                secondpath=null;
                System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
                if (daterange.size() % 2 == 0) {//number of days in between is even
                    for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                        System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                        firstpath = daterange_path + daterange.get(i);
                        secondpath = daterange_path + daterange.get(j);
                        System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                        first_income = getIncomeValueForDateRange(firstpath);
                        second_income = getIncomeValueForDateRange(secondpath);
        //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
        //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                        sum_income=addThoseArrays(first_income,second_income);
        //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                        if(total_income.size()==0)
                            total_income=sum_income;
                        else
                            total_income=addThoseArrays(total_income,sum_income);
                    }
                    System.out.println("Final Total Income for even size array is "+total_income.size()+" and the array is "+total_income);
                }
                if (daterange.size() % 2 == 1) {//number of days in between is odd
                    for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                        System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                        firstpath = daterange_path + daterange.get(i);
                        secondpath = daterange_path + daterange.get(j);
                        System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                        first_income = getIncomeValueForDateRange(firstpath);
                        second_income = getIncomeValueForDateRange(secondpath);
        //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
        //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                        sum_income=addThoseArrays(first_income,second_income);
        //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                        if(total_income.size()==0)
                            total_income=sum_income;
                        else
                            total_income=addThoseArrays(total_income,sum_income);

                        if (k == daterange.size() / 2) {
                            System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                            firstpath = daterange_path + daterange.get(i);
                            System.out.println("Paths for Last location is " + firstpath);
                            first_income = getIncomeValueForDateRange(firstpath);
                            total_income=addThoseArrays(total_income,first_income);
                        }
                    }
                    System.out.println("Date range for "+daterange_product+" Transaction Reference size is " + eln_daterange_transaction_reference.size() + " and the array is " + eln_daterange_transaction_reference);
                    System.out.println("Final Total Income for uneven size array is "+total_income.size()+" and the array is "+total_income);
                }

                //////////////////////////////////////////////////// END OF CODE FOR DOING INCOME

                eln_daterange_income_value = negateThatArray(total_income);

        //        eln_daterange_income_value= negateThatArray(eln_daterange_income_value);
        eln_daterange_unrealised_surplus = negateThatArray(eln_daterange_unrealised_surplus);


        System.out.println("Date range for " + daterange_product + " Transaction Reference size is " + eln_daterange_transaction_reference.size() + " and the array is ");
        //        System.out.println("Book Name size is " + eln_book_name.size() + " and the array is " + eln_book_name);
        //        System.out.println("Market Value size is " + eln_market_value.size() + " and the array is " + eln_market_value);
        System.out.println("Date range for " + daterange_product + " Book Value size is " + eln_book_value.size() + " and the array is ");
        //        System.out.println("Yesterday's Market Value size is " + eln_market_valuet1.size() + " and the array is " + eln_market_valuet1);
        System.out.println("Date range for " + daterange_product + "  Unrealised Surplus Value size is " + eln_daterange_unrealised_surplus.size() + " and the array is ");
        //        createOneBigDateRangeOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for ELN flow - " + timediff(start, end) + " seconds");

    }

    public static boolean isCRNaPREF(String crn_reference,ArrayList<String> fkey,ArrayList<String> subProductType) {
        boolean result = false;
        Instant start = Instant.now();
        for(int i=0;i<fkey.size();i++) {
            if(fkey.get(i).equals(crn_reference)) {

                if (subProductType.get(i).equals("Subordinated Preference Shares") || subProductType.get(i).equals("Senior Preference Shares")) {
                    System.out.println("Checking "+crn_reference+" against "+fkey.get(i)+" and subProductType "+subProductType.get(i));
                    result = true;
                }
            }
        }
        Instant end = Instant.now();
//        System.out.println("Checked "+crn_reference+" in "+timediff(start,end)+" seconds");
        return result;
    }


    @Test
    public static void fetchCRNforDateRange() throws Exception {
        daterange_product = "CRN";
        instruments.add(daterange_product);
        Instant start = Instant.now();
        String filename = "Static Contract Data Report.csv";
        String filename1 = "Data_Trade_CRN_Bound_" + end_date + ".csv";
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + end_date + ".csv";
        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + end_date + ".csv";
        String f3sheetname = "Realized_Cashflows-" + end_date;
        String fn_sheetname = "Static Contract Data Report 201";
        String filename_prefs = "Data_Trade_CRN_Stat";
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        for (String s: files) {
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
            if (s.contains("Data_Trade_CRN_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
            if (s.contains("Data_Trade_CRN_Stat") && s.contains(".csv"))
                filename_prefs = s.substring(s.indexOf("Data_Trade"));

        }
        String identifier = "CRN";
         ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + thatdaysyesterdayDate(end_date) + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //        System.out.println(end_path+"\t"+filename);
        csvToXLSX(end_path, filename);
        //        System.out.println(end_path+"\t"+filename1);
        csvToXLSX(end_path, filename1);
        //        System.out.println(end_path+"\t"+filename2);
        csvToXLSX(end_path, filename2);
        //        System.out.println(end_path+"\t"+filename3);
        csvToXLSX(end_path, filename3);
        csvToXLSX(end_path, filename_prefs);
        //        System.out.println(start_path+"\t"+filename4);
        csvToXLSX(end_minusone_path, filename4);

        crn_daterange_transaction_reference = readPrimaryKey(end_path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5,fkey_datatrade_stat,comparator6;
        //        System.out.println("Transaction count size is " + crn_transaction_reference.size() + " and the list is " + crn_transaction_reference);
        fkey_datatrade = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 16)); //outstanding notional or bookvalue
        comparator2 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 17 )); // accrued intetest native
        comparator3 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column
        comparator5 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        fkey_abprod = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));
        comparator4 = (readColumnData(end_path, filename_prefs.replace(".csv", ".xlsx"), filename_prefs.replace(".csv", ""), 37)); //allInConsideration column from Data_Trade_CRN_Stat
        fkey_datatrade_stat = (readColumnData(end_path, filename_prefs.replace(".csv", ".xlsx"), filename_prefs.replace(".csv", ""), 0)); //transactionReference column from Data_Trade_CRN_Stat
        comparator6 = (readColumnData(end_path, filename_prefs.replace(".csv", ".xlsx"), filename_prefs.replace(".csv", ""), 31)); //subProductType column from Data_Trade_CRN_Stat







//        for(int i=0;i<crn_daterange_transaction_reference.size();i++) {
//            if(isCRNaPREF(crn_daterange_transaction_reference.get(i),fkey_datatrade_stat,comparator6))
//            {
////                System.out.println("You have a stroke and now stutter");
//            }
//        }


        ////// POPULATION OF BOOK VALUE AND ACCRUED INCOME
        for (int i = 0; i < crn_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (crn_daterange_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    if(isCRNaPREF(crn_daterange_transaction_reference.get(i),fkey_datatrade_stat,comparator6))
                        crn_book_value.add(comparator4.get(j));
                    else
                        crn_book_value.add(comparator1.get(j));
                    crn_accrued_income.add(comparator2.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    crn_book_value.add("0.00");
                    crn_accrued_income.add("0.00");
                }
            }
        }

        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < crn_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (crn_daterange_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    crn_market_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    crn_market_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < crn_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (crn_daterange_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    crn_market_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    crn_market_valuet1.add("0.00");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREALISED SURPLUS VALUE
        BigDecimal bookval_bd, accr_income_bd, marval_bd, appvalue_result, unrealised_result, markett1_bd;
        for (int i = 0; i < crn_daterange_transaction_reference.size(); i++) {
            try {
                bookval_bd = new BigDecimal(crn_book_value.get(i));
                accr_income_bd = new BigDecimal(crn_accrued_income.get(i));
                marval_bd = new BigDecimal(crn_market_value.get(i));
                markett1_bd = new BigDecimal(crn_market_valuet1.get(i));
                appvalue_result = marval_bd.subtract(bookval_bd).subtract(accr_income_bd);
                unrealised_result = marval_bd.subtract(markett1_bd);
                crn_appreciation_value.add(String.valueOf(appvalue_result));
//                crn_unreal_surplus.add(String.valueOf(unrealised_result)); ////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                //                crn_unreal_surplus.add(crn_market_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();
                crn_appreciation_value.add("ERROR");
                crn_unreal_surplus.add("ERROR");
            }
        }

        prettyPrint(crn_daterange_transaction_reference,crn_market_value,crn_book_value,crn_accrued_income,crn_appreciation_value);

        //////////////////////////////////////CODE FOR DOING INCOME VALUE
        ArrayList < String > first_income = new ArrayList < > ();
        ArrayList < String > second_income = new ArrayList < > ();
        ArrayList < String > sum_income = new ArrayList < > ();
        ArrayList < String > total_income = new ArrayList < > ();
        String firstpath = null;
        String secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("INCOME VALUE Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
//                System.out.println("First income size is " + first_income.size() + " and the array is " );
//                System.out.println("Second income size is " + second_income.size() + " and the array is " );
//                System.out.println(first_income);
//                System.out.println(second_income);
                sum_income = addThoseArrays(first_income, second_income);
//                                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);
            }
//                        System.out.println("Final Total Income for even size array is "+total_income.size()+" and the array is "+total_income);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("INCOME VALUE Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
//                System.out.println(first_income);
//                System.out.println(second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_income = getIncomeValueForDateRange(firstpath);
                    total_income = addThoseArrays(total_income, first_income);
                }
            }
            //            System.out.println("Transaction count size is " + crn_daterange_transaction_reference.size() + " and the list is " + crn_daterange_transaction_reference);
            //            System.out.println("Final Total Income for uneven size array is "+total_income.size()+" and the array is "+total_income);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING INCOME

        //////////////////////////////////////CODE FOR DOING UNREAL VALUE
        ArrayList < String > first_unreal = new ArrayList < > ();
        ArrayList < String > second_unreal = new ArrayList < > ();
        ArrayList < String > sum_unreal = new ArrayList < > ();
        ArrayList < String > total_unreal = new ArrayList < > ();
         firstpath = null;
         secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("UNREAL SURPLUS Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is " );
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is " );
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is " );
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);
            }
            //            System.out.println("Final Total unreal for even size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("UNREAL SURPLUS Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is " );
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is ");
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                    System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                    total_unreal = addThoseArrays(total_unreal, first_unreal);
                }
            }
            System.out.println("CRN Transaction count size is " + crn_daterange_transaction_reference.size() + " and the list is " );
            System.out.println("Final Total unreal for uneven size array is " + total_unreal.size() + " and the array is " );
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING unreal
        crn_daterange_unrealised_surplus = total_unreal;

        crn_daterange_income_value = total_income;

        crn_daterange_income_value = negateThatArray(crn_daterange_income_value);
        crn_daterange_unrealised_surplus = negateThatArray(crn_daterange_unrealised_surplus);

        System.out.println("Date range for " + daterange_product + " Transaction Reference size is " + crn_daterange_transaction_reference.size() + " and the array is ");
        //        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        System.out.println("Date range for " + daterange_product + " Book Value size is " + crn_book_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Accrued Income Value is " + crn_accrued_income.size() + " and the array is ");
        //        System.out.println("Book Name size is " + crn_book_name.size() + " and the array is " + crn_book_name);
        System.out.println("Date range for " + daterange_product + " Appreciation Value size is " + crn_appreciation_value.size() + " and the array is ");
        //        System.out.println("Yesterday's Market Value size is " + crn_market_valuet1.size() + " and the array is " + crn_market_valuet1);
        System.out.println("Date range for " + daterange_product + " CRN Date Range Realised Cash Flow or Income Value size is " + crn_daterange_income_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " CRN Date Range Unrealised Surplus Value size is " + crn_daterange_unrealised_surplus.size() + " and the array is ");
        //        createOneBigDateRangeOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for CRN flow - " + timediff(start, end) + " seconds");

    }

    @Test
    public static void fetchIRNforDateRange() throws Exception {
        daterange_product = "IRN";

        instruments.add(daterange_product);
        Instant start = Instant.now();
        String filename = "Static Contract Data Report.csv";
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        for (String s: files) 
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        //        filename = "Static Contract Data Report "+date+".csv";
        String filename1 = "Data_Trade_IRN_Bound_" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_IRN_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + end_date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + end_date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "IRN";

        ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + thatdaysyesterdayDate(end_date) + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //        System.out.println(end_path+"\t"+filename);
        csvToXLSX(end_path, filename);
        //        System.out.println(end_path+"\t"+filename1);
        csvToXLSX(end_path, filename1);
        //        System.out.println(end_path+"\t"+filename2);
        csvToXLSX(end_path, filename2);
        //        System.out.println(end_path+"\t"+filename3);
        csvToXLSX(end_path, filename3);
        System.out.println(end_minusone_path + "\t" + filename4);
        csvToXLSX(end_minusone_path, filename4);

        irn_daterange_transaction_reference = readPrimaryKey(end_path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        System.out.println("Book Name size is " + irn_bookname.size() + " and the array is " + irn_bookname);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6, comparator7;
        fkey_datatrade = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 13)); //outstanding notional column
        comparator2 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 14)); //accrued income column
        comparator3 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column
//        comparator4 = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //realised cash flow
        comparator5 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
//        fk_realizedcash = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //fk realised cash
        fkey_abprod = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));

        for (int i = 0; i < irn_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (irn_daterange_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    irn_outstanding_notional.add(comparator1.get(j));
                    irn_book_value.add(comparator1.get(j));
                    irn_accrued_income.add(comparator2.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    irn_outstanding_notional.add("0.00");
                    irn_book_value.add("0.00");
                    irn_accrued_income.add("0.00");
                }
            }
        }
        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < irn_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (irn_daterange_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    irn_market_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    irn_market_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < irn_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (irn_daterange_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    irn_market_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    irn_market_valuet1.add("0.00");
                }
            }
        }

        //BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREALISED SURPLUS VALUE
        BigDecimal out_not_bd, accr_income_bd, marval_bd, appvalue_result, marvalt1_bd, unrealised_result;
        for (int i = 0; i < irn_daterange_transaction_reference.size(); i++) {
            try {
                out_not_bd = new BigDecimal(irn_outstanding_notional.get(i));
                accr_income_bd = new BigDecimal(irn_accrued_income.get(i));
                marval_bd = new BigDecimal(irn_market_value.get(i));
                marvalt1_bd = new BigDecimal(irn_market_valuet1.get(i));
                appvalue_result = marval_bd.subtract(out_not_bd).subtract(accr_income_bd);
                irn_appreciation_value.add(String.valueOf(appvalue_result));
                unrealised_result = marval_bd.subtract(marvalt1_bd);
                irn_unrealised_surplus_value.add(String.valueOf(unrealised_result)); //this is fixed since this is for a date range
                //UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                // UPDATING AGAIN CAUSE VYV SAID WHAT HE SAID ON FRIDAY WAS INCORRECT
                //                irn_unrealised_surplus_value.add(irn_market_value.get(i));
            } catch (Exception e) {
                e.printStackTrace();
                irn_appreciation_value.add("ERROR");
                irn_unrealised_surplus_value.add("ERROR");
            }
        }


        //////////////////////////////////////CODE FOR DOING UNREAL VALUE
        ArrayList < String > first_unreal = new ArrayList < > ();
        ArrayList < String > second_unreal = new ArrayList < > ();
        ArrayList < String > sum_unreal = new ArrayList < > ();
        ArrayList < String > total_unreal = new ArrayList < > ();
        String firstpath = null, secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                //                System.out.println("Sum unreal size is "+sum_unreal.size()+" and the array is "+sum_unreal);
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);
            }
            System.out.println("Final Total unreal for even size array is " + total_unreal.size() + " and the array is ");
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is ");
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                    total_unreal = addThoseArrays(total_unreal, first_unreal);
                }
            }
            //            System.out.println("Transaction count size is " + irn_daterange_transaction_reference.size() + " and the list is " + irn_daterange_transaction_reference);
            //            System.out.println("Final Total unreal for uneven size array is "+total_unreal.size()+" and the array is "+total_unreal);
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING unreal
        irn_daterange_unrealised_surplus = total_unreal;



        //        Thread.sleep(2000);



        //////////////////////////////////////CODE FOR DOING INCOME VALUE
        ArrayList < String > first_income = new ArrayList < > ();
        ArrayList < String > second_income = new ArrayList < > ();
        ArrayList < String > sum_income = new ArrayList < > ();
        ArrayList < String > total_income = new ArrayList < > ();
        ArrayList < String > temp = new ArrayList < > ();
        ArrayList < String > temp1 = new ArrayList < > ();
        firstpath = null;
        secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);
            }
            System.out.println("Final Total Income for even size array is " + total_income.size() + " and the array is ");
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_income = getIncomeValueForDateRange(firstpath);
                    total_income = addThoseArrays(total_income, first_income);
                }
            }

        }

        //        irn_income_value=total_income;
        irn_daterange_income_value = total_income;


        irn_daterange_income_value = negateThatArray(irn_daterange_income_value);
        irn_daterange_unrealised_surplus = negateThatArray(irn_daterange_unrealised_surplus);

        System.out.println("Date range for " + daterange_product + "  Transaction Reference size is " + irn_daterange_transaction_reference.size() + " and the array is ");
        //        System.out.println("Market Value size is " + irn_market_value.size() + " and the array is " + irn_market_value);
        System.out.println("Date range for " + daterange_product + " Book Value size is " + irn_book_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Accrued Income Value is " + irn_accrued_income.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Realised Cash Flow is " + irn_realised_cash_flow.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Book Name size is " + irn_bookname.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Appreciation Value size is " + irn_appreciation_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Income Value size is " + irn_daterange_income_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Unrealised Surplus Value size is " + irn_daterange_unrealised_surplus.size() + " and the array is ");

        //        createOneBigDateRangeOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for IRN flow - " + timediff(start, end) + " seconds");

    }

    @Test
    public static void fetchIRSforDateRange() throws Exception {

        daterange_product = "IRS";
        instruments.add(daterange_product);
        Instant start = Instant.now();
        String filename = "Static Contract Data Report.csv";
        ArrayList < String > files = getFilenamesFromFolder(end_path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        //        filename = "Static Contract Data Report "+date+".csv";
        String filename1 = "Data_Trade_IRS_Bound_" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_IRS_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + end_date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + end_date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "IRS";

        ///previous date files
        end_minusone_path = thatdaysyesterday(end_date);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(end_minusone_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + thatdaysyesterdayDate(end_date) + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //        System.out.println(end_path+"\t"+filename);
        csvToXLSX(end_path, filename);
        //        System.out.println(end_path+"\t"+filename1);
        csvToXLSX(end_path, filename1);
        //        System.out.println(end_path+"\t"+filename2);
        csvToXLSX(end_path, filename2);
        //        System.out.println(end_path+"\t"+filename3);
        csvToXLSX(end_path, filename3);
        //        System.out.println(start_path+"\t"+filename4);
        csvToXLSX(end_minusone_path, filename4);

        irs_daterange_transaction_reference = readPrimaryKey(end_path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6, comparator7;

        fkey_datatrade = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 24));
        comparator2 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 25));
        comparator3 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 28));
        comparator4 = (readColumnData(end_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 29));
        comparator5 = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8));
//        comparator6 = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7));
        comparator7 = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8));
//        fk_realizedcash = (readColumnData(end_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3));
        fkey_abprod = (readColumnData(end_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(end_minusone_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));


        ////// POPULATION OF PON, RON AND PAI, RAI
        for (int i = 0; i < irs_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (irs_daterange_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    irs_pay_outstanding_notional.add(comparator1.get(j));
                    irs_receive_outstanding_notional.add(comparator2.get(j));
                    irs_pay_accrued_income.add(comparator3.get(j));
                    irs_receive_accrued_income.add(comparator4.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    //
                }
            }
        }


        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < irs_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (irs_daterange_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    irs_market_value.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    irs_market_value.add("0.00");
                }
            }
        }


        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < irs_daterange_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (irs_daterange_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    irs_market_valuet1.add(comparator7.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    irs_market_valuet1.add("0.00");
                }
            }
        }


        //BIG DECIMAL CALCULATIONS FOR BOOK VALUE AND ACCRUED INCOME VALUE
        BigDecimal pay_bd, rec_bd, out_notional_result, pai_bd, rai_bd, accrued_income_result;
        for (int i = 0; i < irs_pay_outstanding_notional.size(); i++) {
            try {
                pay_bd = new BigDecimal(irs_pay_outstanding_notional.get(i));
                rec_bd = new BigDecimal(irs_receive_outstanding_notional.get(i));
                pai_bd = new BigDecimal(irs_pay_accrued_income.get(i));
                rai_bd = new BigDecimal(irs_receive_accrued_income.get(i));
                out_notional_result = rec_bd.add(pay_bd);
                irs_book_value.add(String.valueOf(out_notional_result));
                accrued_income_result = rai_bd.add(pai_bd);
                irs_accrued_income_value.add(String.valueOf(accrued_income_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                irs_book_value.add("ERROR");
                irs_accrued_income_value.add("ERROR");
            }
        }

        //BIG DECIMAL CALCULATIONS for APPRECIATION VALUE
        BigDecimal marval_bd, appvalue_result;
        for (int i = 0; i < irs_daterange_transaction_reference.size(); i++) {
            //            System.out.println("Doing this for values "+irs_pay_outstanding_notional.get(i)+"\n"+irs_receive_outstanding_notional.get(i)+"\n"+irs_pay_accrued_income.get(i)+"\n"+irs_receive_accrued_income.get(i)+"\n"+irs_market_value.get(i));
            try {
                pay_bd = new BigDecimal(irs_pay_outstanding_notional.get(i));
                rec_bd = new BigDecimal(irs_receive_outstanding_notional.get(i));
                pai_bd = new BigDecimal(irs_pay_accrued_income.get(i));
                rai_bd = new BigDecimal(irs_receive_accrued_income.get(i));
                marval_bd = new BigDecimal(irs_market_value.get(i));
                //            System.out.println("And after converting, they are "+pay_bd+"\n"+irs_receive_outstanding_notional.get(i)+"\n"+irs_pay_accrued_income.get(i)+"\n"+irs_receive_accrued_income.get(i)+"\n"+irs_market_value.get(i));
                appvalue_result = marval_bd.subtract(pay_bd).subtract(pai_bd).subtract(rec_bd).subtract(rai_bd);
                irs_appreciation_value.add(String.valueOf(appvalue_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                irs_appreciation_value.add("ERROR");
            }
        }


        //        //BIG DECIMAL CALCULATION OF UNREALISED SURPLUS VALUE
        //        BigDecimal marval_today_bd, marval_yesterday_bd, unrealised_result;
        //        for (int i = 0; i < irs_transaction_reference.size(); i++) {
        //            //            System.out.println("Doing this for values "+irs_pay_outstanding_notional.get(i)+"\n"+irs_receive_outstanding_notional.get(i)+"\n"+irs_pay_accrued_income.get(i)+"\n"+irs_receive_accrued_income.get(i)+"\n"+irs_market_value.get(i));
        //            try {
        //                marval_today_bd = new BigDecimal(irs_market_value.get(i));
        //                marval_yesterday_bd = new BigDecimal(irs_market_valuet1.get(i));
        //                unrealised_result = marval_today_bd.subtract(marval_yesterday_bd);
        //                irs_daterange_unrealised_surplus.add(String.valueOf(unrealised_result)); // this is right since it's for a date range
        //                //UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
        ////                irs_unrealised_surplus_value.add(irs_market_value.get(i));
        //            } catch (Exception e) {
        //
        //                //                e.printStackTrace();
        //                irs_daterange_unrealised_surplus.add("ERROR");
        //            }
        //        }




        //        Thread.sleep(2000);



        //////////////////////////////////////CODE FOR DOING UNREAL VALUE
        ArrayList < String > first_unreal = new ArrayList < > ();
        ArrayList < String > second_unreal = new ArrayList < > ();
        ArrayList < String > sum_unreal = new ArrayList < > ();
        ArrayList < String > total_unreal = new ArrayList < > ();
        String firstpath = null, secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is ");
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);
//                prettyPrint(irs_daterange_transaction_reference,first_unreal,second_unreal,sum_unreal,total_unreal);
            }

            System.out.println("Final Total unreal for even size array is " + total_unreal.size() + " and the array is ");
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                second_unreal = getUnrealSurplusValueForDateRange(secondpath);
                System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                System.out.println("Second unreal size is " + second_unreal.size() + " and the array is ");
                sum_unreal = addThoseArrays(first_unreal, second_unreal);
                System.out.println("Sum unreal size is " + sum_unreal.size() + " and the array is ");
                if (total_unreal.size() == 0)
                    total_unreal = sum_unreal;
                else
                    total_unreal = addThoseArrays(total_unreal, sum_unreal);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_unreal = getUnrealSurplusValueForDateRange(firstpath);
                    System.out.println("First unreal size is " + first_unreal.size() + " and the array is ");
                    total_unreal = addThoseArrays(total_unreal, first_unreal);
                }
//                prettyPrint(irs_daterange_transaction_reference,first_unreal,second_unreal,sum_unreal,total_unreal);
            }

            System.out.println("Final Total unreal for uneven size array is " + total_unreal.size() + " and the array is ");
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING unreal
        irs_daterange_unrealised_surplus = total_unreal;


        //////////////////////////////////////CODE FOR DOING INCOME VALUE
        ArrayList < String > first_income = new ArrayList < > ();
        ArrayList < String > second_income = new ArrayList < > ();
        ArrayList < String > sum_income = new ArrayList < > ();
        ArrayList < String > total_income = new ArrayList < > ();
        firstpath = null;
        secondpath = null;
        System.out.println("Array sizes " + daterange.size() + " and the remainder is " + daterange.size() % 2 + " and the half of it is " + daterange.size() / 2);
        if (daterange.size() % 2 == 0) { //number of days in between is even
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);
                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);
            }
            System.out.println("Final Total Income for even size array is " + total_income.size() + " and the array is ");
        }
        if (daterange.size() % 2 == 1) { //number of days in between is odd
            for (int i = 0, j = 1, k = 1; k <= daterange.size() / 2; i++, i++, j++, j++, k++) {
                System.out.println("Round " + k + "\t Date 1 " + daterange.get(i) + " with Date 2 " + daterange.get(j) + " i and j are " + i + "\t" + j);
                firstpath = daterange_path + daterange.get(i);
                secondpath = daterange_path + daterange.get(j);
                System.out.println("Paths for 2 locations are " + firstpath + "\t" + secondpath);

                first_income = getIncomeValueForDateRange(firstpath);
                second_income = getIncomeValueForDateRange(secondpath);
                //                System.out.println("First income size is "+first_income.size()+" and the array is "+first_income);
                //                System.out.println("Second income size is "+second_income.size()+" and the array is "+second_income);
                sum_income = addThoseArrays(first_income, second_income);
                //                System.out.println("Sum Income size is "+sum_income.size()+" and the array is "+sum_income);
                if (total_income.size() == 0)
                    total_income = sum_income;
                else
                    total_income = addThoseArrays(total_income, sum_income);

                if (k == daterange.size() / 2) {
                    System.out.println("Round " + (k + 1) + "\t Date 1 " + daterange.get(i + 2));
                    firstpath = daterange_path + daterange.get(i + 2);
                    System.out.println("Paths for Last location is " + firstpath);
                    first_income = getIncomeValueForDateRange(firstpath);
                    total_income = addThoseArrays(total_income, first_income);
                }
            }
            System.out.println("Final Total Income for uneven size array is " + total_income.size() + " and the array is ");
        }

        //////////////////////////////////////////////////// END OF CODE FOR DOING INCOME

        irs_daterange_income_value = total_income;

        irs_daterange_income_value = negateThatArray(irs_daterange_income_value);
        irs_daterange_unrealised_surplus = negateThatArray(irs_daterange_unrealised_surplus);


        System.out.println("Date range for " + daterange_product + " Transaction Reference size is " + irs_daterange_transaction_reference.size() + " and the array is ");
        //        System.out.println("Outstanding Notional size is " + irs_outstanding_notional.size() + " and the array is " + irs_outstanding_notional);
        //        System.out.println("Market Value size is " + irs_market_value.size() + " and the array is " + irs_market_value);
        System.out.println("Date range for " + daterange_product + " Book Value size is " + irs_book_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Date range for " + daterange_product + " Accrued Income Value is " + irs_accrued_income_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Book Name size is " + irs_book_name.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Appreciation Value size is " + irs_appreciation_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Income Value size is " + irs_daterange_income_value.size() + " and the array is ");
        System.out.println("Date range for " + daterange_product + " Unrealised Surplus Value size is " + irs_daterange_unrealised_surplus.size() + " and the array is ");

        //        createOneBigDateRangeOutputFile();

        Instant end = Instant.now();
        System.out.println("Time taken for IRS flow - " + timediff(start, end) + " seconds");
    }



    public static ArrayList < String > addThoseArrays(ArrayList < String > one, ArrayList < String > two) {
        BigDecimal onebd, twobd, resbd;
        ArrayList < String > res = new ArrayList < > ();
        //        System.out.println("Sizes = "+one.size()+" and "+two.size());
        if (one.size() == two.size())
            for (int i = 0; i < one.size(); i++) {
                //            System.out.println("ATransaction countdding value "+one.get(i)+" and "+two.get(i)+" where both are ");
                onebd = new BigDecimal(one.get(i));
                twobd = new BigDecimal(two.get(i));
                resbd = onebd.add(twobd);
                res.add(String.valueOf(resbd));
            }
        else {
            System.out.println("Arrays have different sizes");
        }
        //        System.out.println("Added. Result size is "+res.size()+" and the array is "+res);
        return res;
    }


    public static ArrayList < String > getCashMIIncomeValueforDateRange(String folder_path) throws Exception {
        ArrayList < String > files = getFilenamesFromFolder(folder_path);
        ArrayList < String > daterange_income_value_method = new ArrayList < > ();

        String filename1 = "Cash_MI_Daily_PLA-2019" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Cash_MI_Daily_PLA") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Cash_MI_Daily_PLA"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + end_date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + end_date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + end_date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "CashMI";

        csvToXLSX(folder_path, filename1);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5;
        comparator1 = (readColumnData(folder_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 15)); //income
//        comparator2 = (readColumnData(folder_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 19)); // book value
//        comparator3 = (readColumnData(folder_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 20)); // accrued income
        comparator4 = (readColumnData(folder_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 4)); // book name
        comparator5 = (readColumnData(folder_path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 3)); // primary key

        //        for (int i = 2; i < comparator5.size(); i++) {
        //            if(cash_daterange_transaction_reference.get(i).toLowerCase().contains(comparator5.get(i).toLowerCase()))
        //                daterange_income_value_method.add(comparator1.get(i));
        //        }


//prettyPrint(cash_daterange_transaction_reference,comparator1,comparator5);
        for (int i = 0, b = 2; i < cash_daterange_transaction_reference.size(); i++, b++) {
            boolean match_found = false;
            for (int j = 0; j < comparator5.size(); j++) {
                if (comparator5.get(j).contains(cash_daterange_transaction_reference.get(i))) {
                    daterange_income_value_method.add(comparator1.get(j));
                    match_found = true;
                }
                if (!match_found && j == comparator5.size() - 1) {
                    daterange_income_value_method.add("0.00");
                }
            }
        }

//        System.out.println("Comparator for product " + daterange_product + " count size is " + comparator1.size() + " and the list is ");
//        System.out.println("Daterange Income Value for product " + daterange_product + " count size is " + daterange_income_value_method.size() + " and the list is ");
//prettyPrint(cash_daterange_transaction_reference,daterange_income_value_method);
        return daterange_income_value_method;
    }


    public static ArrayList < String > getEQFUTIncomeValueForDateRange(String folder_path) throws Exception {

        ArrayList<String> daterange_transaction_reference_method = new ArrayList<>();
        ArrayList<String> daterange_income_value_method = new ArrayList<>();
        ArrayList<String> daterange_outstanding_value_method = new ArrayList<>();
        String identifier = daterange_product;

        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList<String> files = getFilenamesFromFolder(folder_path);
        for (String s : files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String filename1 = "Data_Trade_EQFUT_Bound_" + date + ".csv";
        for (String s : files)
            if (s.contains("Data_Trade_EQFUT_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin 10D Expected CashFlows Report - 2019" + date + ".csv";
        for (String s : files)
            if (s.contains("AB - PROD - FIN - BO Libfin 10D Expected CashFlows Report - 2019") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin 10D";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s : files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        String fn_sheetname = "Static Contract Data Report 201";
        identifier = "EQFUT";


        csvToXLSX(folder_path, filename1);
        csvToXLSX(folder_path, filename);
        csvToXLSX(folder_path, filename2);
        csvToXLSX(folder_path, filename3);
        daterange_transaction_reference_method = eqf_daterange_transaction_reference;
        //        eqf_daterange_transaction_reference = daterange_transaction_reference_method;
        System.out.println("Date Range Transaction for product " + daterange_product + " count size is " + daterange_transaction_reference_method.size() + " and the list is ");

        ArrayList<String> fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5;
        fkey_abprod = (readColumnData(folder_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 3));
        comparator1 = (readColumnData(folder_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 7)); //outstanding income

//        ////// POPULATION OF OUTSTANDING INCOME
//        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
//            boolean match_found = false;
//            for (int j = 0; j < fkey_abprod.size(); j++) {
//                if (daterange_transaction_reference_method.get(i).equals(fkey_abprod.get(j))) {
//                    daterange_outstanding_value_method.add(comparator1.get(j));
//                    match_found = true;
//                }
//                if (!match_found && j == fkey_abprod.size() - 1) {
//                    daterange_outstanding_value_method.add("0.00");
//                }
//            }
//        }
//
//        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
//        BigDecimal outstandingincome_bd, income_bd;
//        int minus1 = -1;
//        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
//            try {
//                outstandingincome_bd = new BigDecimal(daterange_outstanding_value_method.get(i));
//                income_bd = outstandingincome_bd.multiply(new BigDecimal(minus1));
//                daterange_income_value_method.add(String.valueOf(income_bd));
//            } catch (Exception e) {
//                //                e.printStackTrace();
//                daterange_income_value_method.add("0.00");
//            }
//        }


        ///////////////////////////////////////////////////////////////////////////////////////////////////////

        //TRICKY POPULATION OF REALISED CASH FLOW
        ArrayList<Integer> matchingIndices = new ArrayList<>();
        ArrayList<String> irn_dupes = new ArrayList<>();
        ArrayList<Integer> irn_dupes_indices = new ArrayList<>();
        ArrayList<Integer> irn_frequency = new ArrayList<>();
        ArrayList<String> irn_unique_sums = new ArrayList<>();
        ArrayList<String> irn_unique_references = new ArrayList<>();
        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
            boolean done = false, match_found = false;
            for (int a = 0; a <  fkey_abprod.size(); a++) {
                String element = fkey_abprod.get(a);
                if (daterange_transaction_reference_method.get(i).equals(element)) {
                    match_found = true;
                    matchingIndices.add(a);
                    irn_dupes.add(daterange_transaction_reference_method.get(i));
                    //                    System.out.println(Collections.frequency(irn_transaction_reference,irn_transaction_reference.get(i)));
                    int freq = Collections.frequency(fkey_abprod, daterange_transaction_reference_method.get(i));
                    irn_frequency.add(freq);
                    irn_dupes_indices.add(i);
                    if (freq == 1) {
                        irn_unique_sums.add(comparator1.get(a));
                        irn_unique_references.add(fkey_abprod.get(a));
                    }
                    if (freq == 2) {
                        if (!irn_unique_references.contains(fkey_abprod.get(a)))
                            irn_unique_references.add(fkey_abprod.get(a));
                        BigDecimal v, k;
                        v = new BigDecimal(comparator1.get(a));
                        a++;
                        k = new BigDecimal(comparator1.get(a));
                        irn_unique_sums.add(String.valueOf(v.add(k)));
                        //                        System.out.println("Adding v and k "+v+"\t"+k+" and the result is "+v.add(k));
                    }
                    if (freq == 3) { //this loop is not dev-complete
                        if (!irn_unique_references.contains(fkey_abprod.get(a)))
                            irn_unique_references.add(fkey_abprod.get(a));
                        BigDecimal v, k, m;
                        v = new BigDecimal(comparator1.get(a));
                        a++;
                        k = new BigDecimal(comparator1.get(a));
                        a++;
                        m = new BigDecimal(comparator1.get(a));
                        irn_unique_sums.add(String.valueOf(v.add(k).add(m)));


//                        for (int h = 0; h < freq; h++) {
//                            if (h == 0)
//                                irn_unique_sums.add("0.00");
//                        }
                    }

                }
            }
            if (match_found) {
            }
        }

        System.out.println("Matching Indices size " + matchingIndices.size() + " and array is ");
        System.out.println("IRN Dupes size is " + irn_dupes.size() + " and array is ");
        System.out.println("IRN Dupes Indices size is " + irn_dupes_indices.size() + " and array is ");
        System.out.println("IRN Frequency size is " + irn_frequency.size() + " and array is ");
        System.out.println("IRN unique references size is " + irn_unique_references.size() + " and array is ");
        System.out.println("IRN unique sums size is " + irn_unique_sums.size() + " and array is ");
        ArrayList<String> matchingIndices_string = arrayListIntString(matchingIndices);
        ArrayList<String> irn_dupes_indices_string = arrayListIntString(irn_dupes_indices);
        ArrayList<String> irn_frequency_string = arrayListIntString(irn_frequency);
//                prettyPrint(matchingIndices_string,irn_dupes,irn_dupes_indices_string,irn_frequency_string,irn_unique_sums,irn_unique_references);

        //TRICKY RE-POPULATION OF IRN INCOME VALUES
        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
            if (irn_unique_references.contains(daterange_transaction_reference_method.get(i))) {
                int index = irn_unique_references.indexOf(daterange_transaction_reference_method.get(i));
                daterange_income_value_method.add(irn_unique_sums.get(index));
            } else
                daterange_income_value_method.add("0.00");
        }
        //        System.out.println("Yipee Kaay Yaay size is "+daterange_income_value_method.size()+" and the array is "+daterange_income_value_method);

        ////////////////////////////////////////////////////////////////////////////////////////////////

        System.out.println("Income Value size is  " + daterange_income_value_method.size() + " and the array is ");
        prettyPrint(daterange_transaction_reference_method,daterange_income_value_method);
        return daterange_income_value_method;


    }

    public static ArrayList < String > getIncomeValueForDateRange(String folder_path) throws Exception {
        ArrayList < String > daterange_transaction_reference_method = new ArrayList < > ();
        ArrayList < String > daterange_income_value_method = new ArrayList < > ();
        String identifier = daterange_product;
        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(folder_path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String fn_sheetname = "Static Contract Data Report 201";

        String filename3 = null;
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));

        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";

        System.out.println("Filename " + filename);
        System.out.println("Filename 3 " + filename3);
        csvToXLSX(folder_path, filename);
        csvToXLSX(folder_path, filename3);
        csvToXLSX(folder_path, filename2);
        ArrayList < String > comparator4;
        ArrayList < String > fk_realizedcash;
        comparator4 = (readColumnData(folder_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //realised cash flow
        fk_realizedcash = (readColumnData(folder_path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //fk realised cash
        if (daterange_product.contains("CFS"))
            daterange_transaction_reference_method = cfs_daterange_transaction_reference;
        else
            daterange_transaction_reference_method = readPrimaryKey(folder_path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);

        if (daterange_product.equalsIgnoreCase("IRN"))
            daterange_transaction_reference_method = irn_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("CRN"))
            daterange_transaction_reference_method = crn_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("IRS"))
            daterange_transaction_reference_method = irs_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("ELN(A)"))
            daterange_transaction_reference_method = eln_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("EQFUT"))
            daterange_transaction_reference_method = eqf_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("EQFWD"))
            daterange_transaction_reference_method = eqfwd_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("IRFWD"))
            daterange_transaction_reference_method = irfwd_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("CFS"))
            daterange_transaction_reference_method = cfs_daterange_transaction_reference;
        System.out.println("Date Range Transaction for product " + daterange_product + " count size is " + daterange_transaction_reference_method.size() + " and the list is ");

        ///////////////////////////////////////////////////////////////////////////////////////////////////////

        //TRICKY POPULATION OF INCOME VALUE
        ArrayList < Integer > matchingIndices = new ArrayList < > ();
        ArrayList < String > irn_dupes = new ArrayList < > ();
        ArrayList < Integer > irn_dupes_indices = new ArrayList < > ();
        ArrayList < Integer > irn_frequency = new ArrayList < > ();
        ArrayList < String > irn_unique_sums = new ArrayList < > ();
        ArrayList < String > irn_unique_references = new ArrayList < > ();
        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
            boolean done = false, match_found = false;
            for (int a = 0; a < fk_realizedcash.size(); a++) {
                String element = fk_realizedcash.get(a);
                if (daterange_transaction_reference_method.get(i).equals(element)) {
                    match_found = true;
                    matchingIndices.add(a);
                    irn_dupes.add(daterange_transaction_reference_method.get(i));
                    //                    System.out.println(Collections.frequency(irn_transaction_reference,irn_transaction_reference.get(i)));
                    int freq = Collections.frequency(fk_realizedcash, daterange_transaction_reference_method.get(i));
                    irn_frequency.add(freq);
                    irn_dupes_indices.add(i);
                    if (freq == 1) {
                        irn_unique_sums.add(comparator4.get(a));
                        irn_unique_references.add(fk_realizedcash.get(a));
                    }
                    if (freq == 2) {
                        if (!irn_unique_references.contains(fk_realizedcash.get(a)))
                            irn_unique_references.add(fk_realizedcash.get(a));
                        BigDecimal v, k;
                        v = new BigDecimal(comparator4.get(a));
                        a++;
                        k = new BigDecimal(comparator4.get(a));
                        irn_unique_sums.add(String.valueOf(v.add(k)));
                        //                        System.out.println("Adding v and k "+v+"\t"+k+" and the result is "+v.add(k));
                    }
                    if (freq == 3) { //this loop is now dev-complete and tested // update on 20190730
                        if (!irn_unique_references.contains(fk_realizedcash.get(a)))
                            irn_unique_references.add(fk_realizedcash.get(a));
                        BigDecimal v,k,m;
                        v = new BigDecimal(comparator4.get(a));
                        a++;
                        k = new BigDecimal(comparator4.get(a));
                        a++;
                        m = new BigDecimal(comparator4.get(a));
                        irn_unique_sums.add(String.valueOf(v.add(k).add(m)));


//                        for (int h = 0; h < freq; h++) {
//                            if (h == 0)
//                                irn_unique_sums.add("0.00");
//                        }
                    }

                }
            }
            if (match_found) {}
        }

//                System.out.println("Matching Indices size " + matchingIndices.size() + " and array is " );
//                System.out.println("IRN Dupes size is " + irn_dupes.size() + " and array is " );
//                System.out.println("IRN Dupes Indices size is " + irn_dupes_indices.size() + " and array is ");
//                System.out.println("IRN Frequency size is " + irn_frequency.size() + " and array is " );
//                System.out.println("IRN unique references size is " + irn_unique_references.size() + " and array is ");
//                System.out.println("IRN unique sums size is " + irn_unique_sums.size() + " and array is ");
                ArrayList<String> matchingIndices_string = arrayListIntString(matchingIndices);
                ArrayList<String> irn_dupes_indices_string = arrayListIntString(irn_dupes_indices);
                ArrayList<String> irn_frequency_string = arrayListIntString(irn_frequency);
//                prettyPrint(matchingIndices_string,irn_dupes,irn_dupes_indices_string,irn_frequency_string,irn_unique_sums,irn_unique_references);

        //TRICKY RE-POPULATION OF IRN INCOME VALUES
        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
            if (irn_unique_references.contains(daterange_transaction_reference_method.get(i))) {
                int index = irn_unique_references.indexOf(daterange_transaction_reference_method.get(i));
                daterange_income_value_method.add(irn_unique_sums.get(index));
            } else
                daterange_income_value_method.add("0.00");
        }
        //        System.out.println("Income V size is "+daterange_income_value_method.size()+" and the array is "+daterange_income_value_method);

        ////////////////////////////////////////////////////////////////////////////////////////////////

        System.out.println("Income Value size is  " + daterange_income_value_method.size() + " and the array is ");
        return daterange_income_value_method;

    }


    public static ArrayList<String> arrayListIntString(ArrayList<Integer> arrList) {
        ArrayList<String> res=new ArrayList<>();
        for(int i=0;i<arrList.size();i++) {
            res.add(String.valueOf(arrList.get(i)));
        }
        return res;
    }


    public static ArrayList < String > getUnrealSurplusValueForDateRange(String folder_path) throws Exception {

        ArrayList < String > daterange_transaction_reference_method = new ArrayList < > ();
        ArrayList < String > daterange_unrealised_surplus = new ArrayList < > ();
        ArrayList < String > mart = new ArrayList < > ();
        ArrayList < String > mart1 = new ArrayList < > ();

        String identifier = daterange_product;
        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(folder_path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String fn_sheetname = "Static Contract Data Report 201";

        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + start_date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));
        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        ///previous date files
        String previous_day_path = thatdaysyesterday(folder_path.substring(folder_path.length() - 8));
        System.out.println(previous_day_path);
        ArrayList < String > previous_day_files = getFilenamesFromFolder(previous_day_path);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + previous_day_path + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));
        ///// previous date files
        System.out.println("Filename " + filename);
        System.out.println("Filename 3 " + filename2);
        csvToXLSX(folder_path, filename);
        csvToXLSX(folder_path, filename2);
        csvToXLSX(previous_day_path, filename4);
        ArrayList < String > fkey_abprod, fkey_abprod_yesterday, comparator5, comparator7;
        comparator5 = (readColumnData(folder_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8));
        comparator7 = (readColumnData(previous_day_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 8));
        if (daterange_product.contains("CFS"))
            daterange_transaction_reference_method = cfs_daterange_transaction_reference;
        else
            daterange_transaction_reference_method = readPrimaryKey(folder_path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        fkey_abprod = (readColumnData(folder_path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(previous_day_path, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));

        if (daterange_product.equalsIgnoreCase("IRN"))
            daterange_transaction_reference_method = irn_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("CRN"))
            daterange_transaction_reference_method = crn_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("IRS"))
            daterange_transaction_reference_method = irs_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("ELN(A)"))
            daterange_transaction_reference_method = eln_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("EQFUT"))
            daterange_transaction_reference_method = eqf_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("EQFWD"))
            daterange_transaction_reference_method = eqfwd_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("IRFWD"))
            daterange_transaction_reference_method = irfwd_daterange_transaction_reference;
        if (daterange_product.equalsIgnoreCase("CFS"))
            daterange_transaction_reference_method = cfs_daterange_transaction_reference;
        System.out.println("Transaction count size is " + daterange_transaction_reference_method.size() + " and the list is ");

        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (daterange_transaction_reference_method.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    mart1.add(comparator7.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    mart1.add("0.00");
                }
            }
        }

        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (daterange_transaction_reference_method.get(i).equals(fkey_abprod.get(j))) {
                    mart.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    mart.add("0.00");
                }
            }
        }

        //        System.out.println("Market value at folder "+);

        //BIG DECIMAL CALCULATION OF UNREALISED SURPLUS VALUE
        BigDecimal marval_today_bd, marval_yesterday_bd, unrealised_result;
        for (int i = 0; i < daterange_transaction_reference_method.size(); i++) {
            //            System.out.println("Doing this for values "+irs_pay_outstanding_notional.get(i)+"\n"+irs_receive_outstanding_notional.get(i)+"\n"+irs_pay_accrued_income.get(i)+"\n"+irs_receive_accrued_income.get(i)+"\n"+irs_market_value.get(i));
            try {
                marval_today_bd = new BigDecimal(mart.get(i));
                marval_yesterday_bd = new BigDecimal(mart1.get(i));
                unrealised_result = marval_today_bd.subtract(marval_yesterday_bd);
                if(folder_path.contains(start_date))
                    daterange_unrealised_surplus.add(String.valueOf(marval_today_bd));
                else
                    daterange_unrealised_surplus.add(String.valueOf(unrealised_result));
                //UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                // UPDATING AGAIN BECAUSE Vyv SAID WHAT HE SAID ON FRIDAY WAS INCORRECT
                //                irs_unrealised_surplus_value.add(irs_market_value.get(i));

            } catch (Exception e) {

                //                e.printStackTrace();
                daterange_unrealised_surplus.add("ERROR");
            }
        }

        return daterange_unrealised_surplus;
    }

    public static void logIntoExcel(String folder_path, ArrayList<String> primarykey, ArrayList<String> data) throws Exception {
        String filepath = "C:\\Users\\vzk1008\\.jenkins\\workspace\\Daily-Run_Source-to-Template_Date-Range\\excelog.xlsx";
        String filename = "excelog.xlsx";
        String logsheet = "excelog";
        setExcelFile(filepath, logsheet);
        System.out.println("Row count is " + getRowCount(logsheet) + " and column count is " + getColumnCount(logsheet, 0) + "path " + filepath + "\nfilename " + filename + "\nsheet name" + logsheet);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.getSheet(logsheet);
        int rowNum = 0;
        int colNum = getColumnCount(logsheet, 0);
        int free;
        for (free = 0; !getCellData(1, 1, logsheet).isEmpty(); free++) {
        }
        System.out.println("I is "+free);
        for (int i = 1; !getCellData(i, 0, logsheet).isEmpty(); i++) {
            if (!getCellData(i, 1, logsheet).isEmpty()) {
                setCellData(" seconds",i, free, logsheet,filename);
            }
        }
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(colNum++);
        cell.setCellValue(daterange_product);
        cell = row.createCell(colNum++);
        cell.setCellValue(folder_path);
        cell = row.createCell(colNum++);
        cell.setCellValue(Instant.now().toString());
        for(int i=3,j=0;i<primarykey.size();i++,j++) {
            cell = row.createCell(3);
            cell.setCellValue(primarykey.get(j));
        }

    }

    public static void prettyPrint(ArrayList < String > one, ArrayList < String > two) {
        System.out.println("Pretty Printing");
        for (int i = 0; i < Math.max(one.size(), two.size()); i++)
            try {
                System.out.println();
                System.out.print(one.get(i));
                System.out.print("\t" + two.get(i));
            } catch (Exception e) {}
    }

    public static void prettyPrint(ArrayList < String > one, ArrayList < String > two, ArrayList < String > three) {
        System.out.println("Pretty Printing");
        for (int i = 0; i < Math.max(one.size(), two.size()); i++)
            try {
                System.out.println();
                System.out.print(one.get(i));
                System.out.print("\t" + two.get(i));
                System.out.print("\t" + three.get(i));
            } catch (Exception e) {}
    }

    public static void prettyPrint(ArrayList < String > one, ArrayList < String > two, ArrayList < String > three, ArrayList < String > four) {
        System.out.println("Pretty Printing");
        for (int i = 0; i < Math.max(one.size(), two.size()); i++)
            try {
                System.out.println();
                System.out.print(one.get(i));
                System.out.print("\t" + two.get(i));
                System.out.print("\t" + three.get(i));
                System.out.print("\t" + four.get(i));
            } catch (Exception e) {}
    }

    public static void prettyPrint(ArrayList < String > one, ArrayList < String > two, ArrayList < String > three, ArrayList < String > four, ArrayList < String > five) {
        System.out.println("Pretty Printing");
        for (int i = 0; i < Math.max(Math.max(one.size(), two.size()), Math.max(three.size(), four.size())); i++)
            try {
                System.out.println();
                System.out.print(one.get(i));
                System.out.print("\t" + two.get(i));
                System.out.print("\t" + three.get(i));
                System.out.print("\t" + four.get(i));
                System.out.print("\t" + five.get(i));
            } catch (Exception e) {}
    }

    public static void prettyPrint(ArrayList < String > one, ArrayList < String > two, ArrayList < String > three, ArrayList < String > four, ArrayList < String > five,ArrayList <String> six) {
        System.out.println("Pretty Printing");
        for (int i = 0; i < Math.max(Math.max(one.size(), two.size()), Math.max(three.size(), four.size())); i++)
            try {
                System.out.println();
                System.out.print(one.get(i));
                System.out.print("\t" + two.get(i));
                System.out.print("\t" + three.get(i));
                System.out.print("\t" + four.get(i));
                System.out.print("\t" + five.get(i));
                System.out.print("\t" + six.get(i));
            } catch (Exception e) {}
    }

    public static void calculateDays() {
        for (int i = 0; i < 1000; i++) {
            daysgone.add(LocalDate.now().minusDays(i).toString().replace("-", ""));
        }
        //        System.out.println(daysgone);
    }

    //    @BeforeSuite
    public static void beforeSuite() {
        path = path + date;
        tempath = path;
        pathminusone = pathminusone + dateminusone;
        ArrayList < String > files = getFilenamesFromFolder(path);
        String filename = "Static Contract Data Report 20190402.csv";
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));
        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String fn_sheetname = "Static Contract Data Report 201";
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));
        csvToXLSX(path, filename);
        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(pathminusone, filename4);
    }

    @Test
    public static void EQFUTFlow() throws Exception {
        //        path = path + date;
        //        tempath = path;
        //        pathminusone = pathminusone + dateminusone;
        product = "EQF";
        instruments.add(product);
        Instant start = Instant.now();
        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String filename1 = "Data_Trade_EQFUT_Bound_" + date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_EQFUT_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin 10D Expected CashFlows Report - 2019" + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin 10D Expected CashFlows Report - 2019") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin 10D";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "EQFUT";

        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));


        csvToXLSX(path, filename1);
        csvToXLSX(path, filename);
        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(pathminusone, filename4);

        System.out.println("reading primary key from file " + filename + "at path " + path + " and sheetname " + fn_sheetname + " with identifier " + identifier);
        eqf_transaction_reference = readPrimaryKey(path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5;
        System.out.println("Transaction count size is " + eqf_transaction_reference.size() + " and the list is " + eqf_transaction_reference);
        //        fkey_datatrade = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 0));
        comparator1 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 7)); //outstanding income
        //        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 17)); // accrued intetest native
        //        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column//i column in ab prod
        //        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //
        //        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        //        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 3));
        //        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));

        //        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        //        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        //        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        //        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        ////// POPULATION OF OUTSTANDING INCOME
        for (int i = 0; i < eqf_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (eqf_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    eqf_outstanding_income.add(comparator1.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    eqf_outstanding_income.add("0.00");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        BigDecimal outstandingincome_bd, income_bd;
        int minus1 = -1;
        for (int i = 0; i < eqf_transaction_reference.size(); i++) {
            try {
                outstandingincome_bd = new BigDecimal(eqf_outstanding_income.get(i));
                income_bd = outstandingincome_bd.multiply(new BigDecimal(minus1));
                eqf_income.add(String.valueOf(income_bd));
            } catch (Exception e) {
                //                e.printStackTrace();
                eqf_income.add("0.00");
            }
        }

        System.out.println("Transaction reference  size is " + eqf_transaction_reference.size() + " and the array is " + eqf_transaction_reference);
        //        System.out.println("Book Name sizeis " + eqf_book_name.size() + " and the array is " + eqf_book_name);
        System.out.println("Outstanding income size is " + eqf_outstanding_income.size() + " and the array is " + eqf_outstanding_income);
        System.out.println("Income size is " + eqf_income.size() + " and the array is " + eqf_income);
        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for EQF flow - " + timediff(start, end) + " seconds");

    }

    @Test
    public static void CashMIFlow() throws Exception {
        //        path = path + date;
        //        tempath = path;createOutput
        //        pathminusone = pathminusone + dateminusone;
        product = "CashMI";
        instruments.add(product);
        Instant start = Instant.now();

        ArrayList < String > files = getFilenamesFromFolder(path);
        //        String filename = "Static Contract Data Report 20190402.csv";
        //        for (String s: files)
        //            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
        //                filename = s.substring(s.indexOf("Static"));
        String filename1 = "Cash_MI_Daily_PLA-2019" + date + ".csv";
        for (String s: files)
            if (s.contains("Cash_MI_Daily_PLA-2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Cash_MI_Daily_PLA"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "CashMI";

        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //book value = 19
        //accrued income = 20
        //income = 15
        csvToXLSX(path, filename1);
        //        csvToXLSX(path, filename);
        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(pathminusone, filename4);


        //        cash_transaction_reference = readPrimaryKey(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5;

        //        cash_transaction_reference = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 3));
        //        System.out.println("Transaction count size is " + cash_transaction_reference.size() + " and the list is " + cash_transaction_reference);
        comparator1 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 15)); //income
        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 19)); // book value
        comparator3 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 20)); // accrued income
        //        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column//i column in ab prod
        comparator4 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 4)); // book name
        comparator5 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 3)); // primary key
        //        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        //        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //
        //        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        //        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));

        //        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        //        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        //        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        //        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        for (int i = 2; i < comparator5.size(); i++) {
            cash_transaction_reference.add(comparator5.get(i));
            cash_book_name.add(comparator4.get(i));
            cash_book_value.add(comparator2.get(i));
            cash_accrued_income.add(comparator3.get(i));
            cash_income.add(comparator1.get(i));
        }

        for (String s: cash_transaction_reference) {
            String n = s.replace("\"", "");
            cash_transaction_reference_new.add(n);
        }

        for (String s: cash_book_name) {
            String n = s.replace("\"", "");
            cash_book_name_new.add(n);
        }

        cash_income = negateThatArray(cash_income);

        System.out.println("Transaction Reference size is " + cash_transaction_reference.size() + " and the array is " + cash_transaction_reference);
        //        System.out.println("Book Name size is " + cash_book_name.size() + " and the array is " + cash_book_name);
        System.out.println("Accrued Income size is " + cash_accrued_income.size() + " and the array is " + cash_accrued_income);
        System.out.println("Book Value size is " + cash_book_value.size() + " and the array is " + cash_book_value);
        System.out.println("Income size is " + cash_income.size() + " and the array is " + cash_income);




        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for CashMI flow - " + timediff(start, end) + " seconds");

    }

    @Test
    public static void EQFWDFlow() throws Exception {
        product = "EQFWD";
        instruments.add(product);
        Instant start = Instant.now();
        ArrayList < String > files = getFilenamesFromFolder(path);
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(pathminusone, filename4);

        ArrayList < String > fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6;

        comparator1 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //primary key
        comparator2 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 1)); //bookname
        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //appreciation value//i column in ab prod
        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //income value//multiply this with -1
        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's appreciation value
        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //foreign key in realised cashflows file
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //fk in ab prod file//keep this as backup
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5)); //
        comparator6 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //

        for (int i = 0; i < comparator1.size(); i++) { //filter primary keys
            if (comparator1.get(i).contains("EQFWD")) {
                eqfwd_transaction_reference.add(comparator1.get(i));
            }
        }


        ////// POPULATION OF BOOK VALUE and APPRECIATION VALUE
        for (int i = 0; i < eqfwd_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < comparator1.size(); j++) {
                if (eqfwd_transaction_reference.get(i).equals(comparator1.get(j))) {
                    eqfwd_book_name.add(comparator2.get(j));
                    eqfwd_appreciation_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == comparator1.size() - 1) {
                    eqfwd_book_name.add("0.00");
                    eqfwd_appreciation_value.add("0.00");
                }
            }
        }

        //YESTERDAY'S APPRECIATION VALUE
        for (int i = 0; i < eqfwd_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (eqfwd_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    eqfwd_appreciation_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    eqfwd_appreciation_valuet1.add("0.00");
                }
            }
        }

        ////// POPULATION OF INCOME VALUE
        for (int i = 0; i < eqfwd_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fk_realizedcash.size(); j++) {
                if (eqfwd_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
                    eqfwd_income_value.add(comparator4.get(j));
                    match_found = true;
                }
                if (!match_found && j == fk_realizedcash.size() - 1) {
                    eqfwd_income_value.add("0.00");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        BigDecimal appvalue_bd, appvaluet1_bd, unrealised_result;
        for (int i = 0; i < eqfwd_transaction_reference.size(); i++) {
            try {
                appvalue_bd = new BigDecimal(eqfwd_appreciation_value.get(i));
                appvaluet1_bd = new BigDecimal(eqfwd_appreciation_valuet1.get(i));
                unrealised_result = appvalue_bd.subtract(appvaluet1_bd);
//                eqfwd_unrealised_surplus.add(String.valueOf(unrealised_result)); ////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                                    eqfwd_unrealised_surplus.add(eqfwd_appreciation_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();
                eqfwd_unrealised_surplus.add("ERROR");
            }
        }

        eqfwd_income_value = negateThatArray(eqfwd_income_value);
        eqfwd_unrealised_surplus = negateThatArray(eqfwd_unrealised_surplus);


        System.out.println("Transaction count size is " + eqfwd_transaction_reference.size() + " and the list is " + eqfwd_transaction_reference);
        //        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        //        System.out.println("Book Value size is " + crn_book_value.size() + " and the array is " + crn_book_value);
        //        System.out.println("Accrued Income Value is " + crn_accrued_income.size() + " and the array is " + crn_accrued_income);
        System.out.println("Book Name size is " + eqfwd_book_name.size() + " and the array is " + eqfwd_book_name);
        System.out.println("Appreciation Value size is " + eqfwd_appreciation_value.size() + " and the array is " + eqfwd_appreciation_value);
        System.out.println("Income Value " + eqfwd_income_value.size() + " and the array is " + eqfwd_income_value);
        //        System.out.println("Realised Cash Flow or Income Value size is " + crn_realised_cash_flow.size() + " and the array is " + crn_realised_cash_flow);
        System.out.println("Unrealised Surplus Value size is " + eqfwd_unrealised_surplus.size() + " and the array is " + eqfwd_unrealised_surplus);

        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for EQFWD flow - " + timediff(start, end) + " seconds");


    }

    @Test
    public static void IRFWDFlow() throws Exception {
        product = "IRFWD";
        instruments.add(product);
        Instant start = Instant.now();
        ArrayList < String > files = getFilenamesFromFolder(path);
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(pathminusone, filename4);

        ArrayList < String > fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6;

        comparator1 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //primary key
        comparator2 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 1)); //bookname
        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //appreciation value//i column in ab prod
        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //income value//multiply this with -1
        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's appreciation value
        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //foreign key in realised cashflows file
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //fk in ab prod file//keep this as backup
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5)); //
        comparator6 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //

        for (int i = 0; i < comparator1.size(); i++) { //filter primary keys
            if (comparator1.get(i).contains("IRFWD")) {
                irfwd_transaction_reference.add(comparator1.get(i));
            }
        }


        ////// POPULATION OF BOOK VALUE and APPRECIATION VALUE
        for (int i = 0; i < irfwd_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < comparator1.size(); j++) {
                if (irfwd_transaction_reference.get(i).equals(comparator1.get(j))) {
                    irfwd_book_name.add(comparator2.get(j));
                    irfwd_appreciation_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == comparator1.size() - 1) {
                    irfwd_book_name.add("0.00");
                    irfwd_appreciation_value.add("0.00");
                }
            }
        }

        //YESTERDAY'S APPRECIATION VALUE
        for (int i = 0; i < irfwd_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (irfwd_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    irfwd_appreciation_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    irfwd_appreciation_valuet1.add("0.00");
                }
            }
        }

        ////// POPULATION OF INCOME VALUE
        for (int i = 0; i < irfwd_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fk_realizedcash.size(); j++) {
                if (irfwd_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
                    irfwd_income_value.add(comparator4.get(j));
                    match_found = true;
                }
                if (!match_found && j == fk_realizedcash.size() - 1) {
                    irfwd_income_value.add("0.00");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        BigDecimal appvalue_bd, appvaluet1_bd, unrealised_result;
        for (int i = 0; i < irfwd_transaction_reference.size(); i++) {
            try {
                appvalue_bd = new BigDecimal(irfwd_appreciation_value.get(i));
                appvaluet1_bd = new BigDecimal(irfwd_appreciation_valuet1.get(i));
                unrealised_result = appvalue_bd.subtract(appvaluet1_bd);
//                irfwd_unrealised_surplus.add(String.valueOf(unrealised_result)); //UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                                irfwd_unrealised_surplus.add(irfwd_appreciation_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();
                irfwd_unrealised_surplus.add("ERROR");
            }
        }

        irfwd_income_value = negateThatArray(irfwd_income_value);
        irfwd_unrealised_surplus = negateThatArray(irfwd_unrealised_surplus);


        System.out.println("Transaction count size is " + irfwd_transaction_reference.size() + " and the list is " + irfwd_transaction_reference);
        //        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        //        System.out.println("Book Value size is " + crn_book_value.size() + " and the array is " + crn_book_value);
        //        System.out.println("Accrued Income Value is " + crn_accrued_income.size() + " and the array is " + crn_accrued_income);
        System.out.println("Book Name size is " + irfwd_book_name.size() + " and the array is " + irfwd_book_name);
        System.out.println("Appreciation Value size is " + irfwd_appreciation_value.size() + " and the array is " + irfwd_appreciation_value);
        System.out.println("Income Value " + irfwd_income_value.size() + " and the array is " + irfwd_income_value);
        //        System.out.println("Realised Cash Flow or Income Value size is " + crn_realised_cash_flow.size() + " and the array is " + crn_realised_cash_flow);
        System.out.println("Unrealised Surplus Value size is " + irfwd_unrealised_surplus.size() + " and the array is " + irfwd_unrealised_surplus);

        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for IRFWD flow - " + timediff(start, end) + " seconds");


    }

    @Test
    public static void CFSFlow() throws Exception {
        product = "CFS";
        instruments.add(product);
        Instant start = Instant.now();
        ArrayList < String > files = getFilenamesFromFolder(path);
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(pathminusone, filename4);

        ArrayList < String > fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6;

        comparator1 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //primary key
        comparator2 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 1)); //bookname
        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //appreciation value//i column in ab prod
        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //income value//multiply this with -1
        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's appreciation value
        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //foreign key in realised cashflows file
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5)); //fk in ab prod file//keep this as backup
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5)); //
        comparator6 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //

        for (int i = 0; i < comparator1.size(); i++) { //filter primary keys
            if (comparator1.get(i).contains("CFS")) {
                cfs_transaction_reference.add(comparator1.get(i));
            }
        }


        ////// POPULATION OF BOOK VALUE and APPRECIATION VALUE
        for (int i = 0; i < cfs_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < comparator1.size(); j++) {
                if (cfs_transaction_reference.get(i).equals(comparator1.get(j))) {
                    cfs_book_name.add(comparator2.get(j));
                    cfs_appreciation_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == comparator1.size() - 1) {
                    cfs_book_name.add("0.00");
                    cfs_appreciation_value.add("0.00");
                }
            }
        }

        //YESTERDAY'S APPRECIATION VALUE
        for (int i = 0; i < cfs_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (cfs_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    cfs_appreciation_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    cfs_appreciation_valuet1.add("0.00");
                }
            }
        }

        ////// POPULATION OF INCOME VALUE
        for (int i = 0; i < cfs_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fk_realizedcash.size(); j++) {
                if (cfs_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
                    cfs_income_value.add(comparator4.get(j));
                    match_found = true;
                }
                if (!match_found && j == fk_realizedcash.size() - 1) {
                    cfs_income_value.add("0.00");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        BigDecimal appvalue_bd, appvaluet1_bd, unrealised_result;
        for (int i = 0; i < cfs_transaction_reference.size(); i++) {
            try {
                appvalue_bd = new BigDecimal(cfs_appreciation_value.get(i));
                appvaluet1_bd = new BigDecimal(cfs_appreciation_valuet1.get(i));
                unrealised_result = appvalue_bd.subtract(appvaluet1_bd);
//                cfs_unrealised_surplus.add(String.valueOf(unrealised_result)); ////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                                cfs_unrealised_surplus.add(cfs_appreciation_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();
                cfs_unrealised_surplus.add("ERROR");
            }
        }

        cfs_income_value = negateThatArray(cfs_income_value);
        cfs_unrealised_surplus = negateThatArray(cfs_unrealised_surplus);


        System.out.println("Transaction count size is " + cfs_transaction_reference.size() + " and the list is " + cfs_transaction_reference);
        //        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        //        System.out.println("Book Value size is " + crn_book_value.size() + " and the array is " + crn_book_value);
        //        System.out.println("Accrued Income Value is " + crn_accrued_income.size() + " and the array is " + crn_accrued_income);
        System.out.println("Book Name size is " + cfs_book_name.size() + " and the array is " + cfs_book_name);
        System.out.println("Appreciation Value size is " + cfs_appreciation_value.size() + " and the array is " + cfs_appreciation_value);
        System.out.println("Income Value " + cfs_income_value.size() + " and the array is " + cfs_income_value);
        //        System.out.println("Realised Cash Flow or Income Value size is " + crn_realised_cash_flow.size() + " and the array is " + crn_realised_cash_flow);
        System.out.println("Unrealised Surplus Value size is " + cfs_unrealised_surplus.size() + " and the array is " + cfs_unrealised_surplus);

        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for CFS flow - " + timediff(start, end) + " seconds");


    }

    @Test
    public static void ELNFlow() throws Exception {
        //        path = path + date;
        //        tempath = path;
        //        pathminusone = pathminusone + dateminusone;
        product = "ELN";
        instruments.add(product);
        Instant start = Instant.now();

        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String filename1 = "Data_Trade_ELN(A)_Stat_" + date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_ELN(A)_Stat_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "ELN(A)";

        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        csvToXLSX(path, filename1);


        eln_transaction_reference = readPrimaryKey(path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5;
        fkey_datatrade = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 7)); //outstanding notional or bookvalue
        //        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 17)); // accrued intetest native
        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column//i column in ab prod
        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //
        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        //        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));

        //        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        //        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        //        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        //        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        ////// POPULATION OF BOOK VALUE
        for (int i = 0; i < eln_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (eln_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    eln_book_value.add(comparator1.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    eln_book_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < eln_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (eln_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    eln_market_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    eln_market_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < eln_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (eln_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    eln_market_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    eln_market_valuet1.add("0.00");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREAL SURPLUS
        BigDecimal marval_bd, unrealised_result, markett1_bd, bookval_bd, appvalue_result;
        for (int i = 0; i < eln_transaction_reference.size(); i++) {
            try {
                marval_bd = new BigDecimal(eln_market_value.get(i));
                markett1_bd = new BigDecimal(eln_market_valuet1.get(i));
                bookval_bd = new BigDecimal(eln_book_value.get(i));
                unrealised_result = marval_bd.subtract(markett1_bd);
                appvalue_result = marval_bd.subtract(bookval_bd);
                eln_appreciation_value.add(String.valueOf(appvalue_result));
//                eln_unreal_surplus.add(String.valueOf(unrealised_result)); ////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                                eln_unreal_surplus.add(eln_market_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();
                eln_unreal_surplus.add("ERROR");
            }
        }


        eln_unreal_surplus = negateThatArray(eln_unreal_surplus);


        System.out.println("Transaction Reference size is " + eln_daterange_transaction_reference.size() + " and the array is " + eln_daterange_transaction_reference);
        System.out.println("Book Name size is " + eln_book_name.size() + " and the array is " + eln_book_name);
        //        System.out.println("Market Value size is " + eln_market_value.size() + " and the array is " + eln_market_value);
        System.out.println("Book Value size is " + eln_book_value.size() + " and the array is " + eln_book_value);
        //        System.out.println("Yesterday's Market Value size is " + eln_market_valuet1.size() + " and the array is " + eln_market_valuet1);
        System.out.println("Unrealised Surplus Value size is " + eln_unreal_surplus.size() + " and the array is " + eln_unreal_surplus);
        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for ELN flow - " + timediff(start, end) + " seconds");





    }

    @Test
    public static void CRNFlow() throws Exception {
        //        path = path + date;
        //        tempath = path;
        //        pathminusone = pathminusone + dateminusone;
        product = "CRN";
        instruments.add(product);
        Instant start = Instant.now();

        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String filename1 = "Data_Trade_CRN_Bound_" + date + ".csv";
        String filename_prefs = "Data_Trade_CRN_Stat";
        for (String s: files) {
            if (s.contains("Data_Trade_CRN_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
            if (s.contains("Data_Trade_CRN_Stat") && s.contains(".csv"))
                filename_prefs = s.substring(s.indexOf("Data_Trade"));
        }
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "CRN";

        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        csvToXLSX(path, filename);
        csvToXLSX(path, filename1);
        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(path,filename_prefs);

        csvToXLSX(pathminusone, filename4);

        crn_transaction_reference = readPrimaryKey(path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5,fkey_datatrade_stat,comparator6,comparator7;
        System.out.println("Transaction count size is " + crn_transaction_reference.size() + " and the list is " + crn_transaction_reference);
        fkey_datatrade = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 16)); //outstanding notional or bookvalue
        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 17)); // accrued intetest native
        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column
        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //
        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));
        comparator6 = (readColumnData(path, filename_prefs.replace(".csv", ".xlsx"), filename_prefs.replace(".csv", ""), 37)); //allInConsideration column from Data_Trade_CRN_Stat
        fkey_datatrade_stat = (readColumnData(path, filename_prefs.replace(".csv", ".xlsx"), filename_prefs.replace(".csv", ""), 0)); //transactionReference column from Data_Trade_CRN_Stat
        comparator7 = (readColumnData(path, filename_prefs.replace(".csv", ".xlsx"), filename_prefs.replace(".csv", ""), 31)); //subProductType column from Data_Trade_CRN_Stat

        //        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        //        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        System.out.println("Fkey Realised size is " + fk_realizedcash.size() + " and the list is " + fk_realizedcash);
        //        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        //        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        ////// POPULATION OF BOOK VALUE AND ACCRUED INCOME
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (crn_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    if(isCRNaPREF(crn_transaction_reference.get(i),fkey_datatrade_stat,comparator7))
                        crn_book_value.add(comparator6.get(j));
                    else
                    crn_book_value.add(comparator1.get(j));
                    crn_accrued_income.add(comparator2.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    crn_book_value.add("0.00");
                    crn_accrued_income.add("0.00");
                }
            }
        }

        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (crn_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    crn_market_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    crn_market_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (crn_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    crn_market_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    crn_market_valuet1.add("0.00");
                }
            }
        }

        //        ////// POPULATION OF REALISED CASH FLOW
        //        for (int i = 0; i < crn_transaction_reference.size(); i++) {
        //            boolean match_found = false;
        //            int j;
        //            for (j = 0; j < fk_realizedcash.size(); j++) {
        //                if (crn_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
        //                    crn_realised_cash_flow.add(comparator4.get(j));
        //                    match_found = true;
        //                    if(match_found)
        //                    {System.out.println("Found at i = "+i+" and j = "+j+" Conbo is "+crn_transaction_reference.get(i)+"\t"+fk_realizedcash.get(j)+"\t"+comparator4.get(j));
        //                    }
        ////                    if(i>100)
        ////                        System.exit(1);
        //                    break;
        //                }
        //                if (j == fk_realizedcash.size() - 1 && !match_found) {
        //                    crn_realised_cash_flow.add("0.00");
        //                    System.out.println("Found at i = "+i+" and j = "+j+" Conbo is "+crn_transaction_reference.get(i)+"\t"+fk_realizedcash.get(j)+"\t"+comparator4.get(j));
        ////                    Thread.sleep(111);
        //                }
        //            }
        //
        //        }


        //TRICKY POPULATION OF REALISED CASH FLOW
        ArrayList < Integer > matchingIndices = new ArrayList < Integer > ();
        ArrayList < String > crn_dupes = new ArrayList < String > ();
        ArrayList < Integer > crn_dupes_indices = new ArrayList < Integer > ();
        ArrayList < Integer > crn_frequency = new ArrayList < Integer > ();
        ArrayList < String > crn_unique_sums = new ArrayList < String > ();
        ArrayList < String > crn_unique_references = new ArrayList < String > ();
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            boolean done = false, match_found = false;
            for (int a = 0; a < fk_realizedcash.size(); a++) {
                String element = fk_realizedcash.get(a);
                if (crn_transaction_reference.get(i).equals(element)) {
                    match_found = true;
                    matchingIndices.add(a);
                    crn_dupes.add(crn_transaction_reference.get(i));
                    //                    System.out.println(Collections.frequency(crn_transaction_reference,crn_transaction_reference.get(i)));
                    int freq = Collections.frequency(fk_realizedcash, crn_transaction_reference.get(i));
                    crn_frequency.add(freq);
                    crn_dupes_indices.add(i);
                    if (freq == 1) {
                        crn_unique_sums.add(comparator4.get(a));
                        crn_unique_references.add(fk_realizedcash.get(a));
                    }
                    if (freq == 2) {
                        if (!crn_unique_references.contains(fk_realizedcash.get(a)))
                            crn_unique_references.add(fk_realizedcash.get(a));
                        BigDecimal v, k;
                        v = new BigDecimal(comparator4.get(a));
                        a++;
                        k = new BigDecimal(comparator4.get(a));
                        crn_unique_sums.add(String.valueOf(v.add(k)));
                        System.out.println("Adding v and k " + v + "\t" + k + " and the result is " + v.add(k));
                        //                        for (int h = 0; h < freq; h++) {
                        //                            if (h == 0)
                        //                                crn_unique_sums.add("puski");
                        //                        }

                    }
                    if (freq == 3) { //this loop is not dev-complete
                        if (!crn_unique_references.contains(fk_realizedcash.get(a)))
                            crn_unique_references.add(fk_realizedcash.get(a));
                        for (int h = 0; h < freq; h++) {
                            if (h == 0)
                                crn_unique_sums.add("puski");
                        }
                    }

                }
            }
            if (match_found) {}
            //                System.out.println("For "+i+" the count is "+matchingIndices);
            //            Thread.sleep(200);
            //            if(matchingIndices.size()>1)
            //                System.exit(1);

            //            if(i>10000)
            //                System.exit(1);
        }

//        System.out.println("Matching Indices size " + matchingIndices.size() + " and array is " + matchingIndices);
//        System.out.println("CRN Dupes size is " + crn_dupes.size() + " and array is " + crn_dupes);
//        System.out.println("CRN Dupes Indices size is " + crn_dupes_indices.size() + " and array is " + crn_dupes_indices);
//        System.out.println("CRN Frequency size is " + crn_frequency.size() + " and array is " + crn_frequency);
//        System.out.println("CRN unique references size is " + crn_unique_references.size() + " and array is " + crn_unique_references);
//        System.out.println("CRN unique sums size is " + crn_unique_sums.size() + " and array is " + crn_unique_sums);

        //TRICKY RE-POPULATION OF CRN INCOME VALUES
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            if (crn_unique_references.contains(crn_transaction_reference.get(i))) {
                int index = crn_unique_references.indexOf(crn_transaction_reference.get(i));
                crn_realised_cash_flow.add(crn_unique_sums.get(index));
            } else
                crn_realised_cash_flow.add("0.00");
        }
        System.out.println("Yipee Kaay Yaay size is " + crn_realised_cash_flow.size() + " and the array is " + crn_realised_cash_flow);





        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREALISED SURPLUS VALUE
        BigDecimal bookval_bd, accr_income_bd, marval_bd, appvalue_result, unrealised_result, markett1_bd;
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            try {
                bookval_bd = new BigDecimal(crn_book_value.get(i));
                accr_income_bd = new BigDecimal(crn_accrued_income.get(i));
                marval_bd = new BigDecimal(crn_market_value.get(i));
                markett1_bd = new BigDecimal(crn_market_valuet1.get(i));
                appvalue_result = marval_bd.subtract(bookval_bd).subtract(accr_income_bd);
                unrealised_result = marval_bd.subtract(markett1_bd);
                crn_appreciation_value.add(String.valueOf(appvalue_result));
//                crn_unreal_surplus.add(String.valueOf(unrealised_result)); ////UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                                crn_unreal_surplus.add(crn_market_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();
                crn_appreciation_value.add("ERROR");
                crn_unreal_surplus.add("ERROR");
            }
        }

        crn_unreal_surplus = negateThatArray(crn_unreal_surplus);
        crn_realised_cash_flow = negateThatArray(crn_realised_cash_flow);

        System.out.println("Transaction Reference size is " + crn_transaction_reference.size() + " and the array is " + crn_transaction_reference);
        //        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        System.out.println("Book Value size is " + crn_book_value.size() + " and the array is " + crn_book_value);
        System.out.println("Accrued Income Value is " + crn_accrued_income.size() + " and the array is " + crn_accrued_income);
        //        System.out.println("Book Name size is " + crn_book_name.size() + " and the array is " + crn_book_name);
        System.out.println("Appreciation Value size is " + crn_appreciation_value.size() + " and the array is " + crn_appreciation_value);
        //        System.out.println("Yesterday's Market Value size is " + crn_market_valuet1.size() + " and the array is " + crn_market_valuet1);
        System.out.println("Realised Cash Flow or Income Value size is " + crn_realised_cash_flow.size() + " and the array is " );
        System.out.println("Unrealised Surplus Value size is " + crn_unreal_surplus.size() + " and the array is " + crn_unreal_surplus);
        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for CRN flow - " + timediff(start, end) + " seconds");

    }

    @Test
    public static void IRNFlow() throws Exception {
        //        path = path + date;
        //        tempath = path;
        //        pathminusone = pathminusone + dateminusone;
        product = "IRN";
        instruments.add(product);
        Instant start = Instant.now();
        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        //        filename = "Static Contract Data Report "+date+".csv";
        String filename1 = "Data_Trade_IRN_Bound_" + date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_IRN_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "IRN";


        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        //        csvToXLSX(path, filename);
        csvToXLSX(path, filename1);
        //        csvToXLSX(path, filename2);
        //        csvToXLSX(path, filename3);
        //        csvToXLSX(pathminusone, filename4);

        irn_transaction_reference = readPrimaryKey(path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6, comparator7;
        System.out.println(irn_transaction_reference);
        System.out.println("Transaction count size is " + irn_transaction_reference.size() + " and the list is " + irn_transaction_reference);
        fkey_datatrade = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 13)); //outstanding notional column
        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 14)); //accrued income column
        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column
        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //realised cash flow
        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //fk realised cash
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));


        //        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        //        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        //        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        //        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        ////// POPULATION OF BOOK VALUE AND OUTSTANDING NOTIONAL AND ACCRUED INCOME
        for (int i = 0; i < irn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (irn_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    irn_outstanding_notional.add(comparator1.get(j));
                    irn_book_value.add(comparator1.get(j));
                    irn_accrued_income.add(comparator2.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    irn_outstanding_notional.add("0.00");
                    irn_book_value.add("0.00");
                    irn_accrued_income.add("0.00");
                }
            }
        }
        //IRS1306713

        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < irn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (irn_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    irn_market_value.add(comparator3.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    irn_market_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF REALISED CASH FLOW AND INCOME VALUE
        for (int i = 0; i < irn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fk_realizedcash.size(); j++) {
                if (irn_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
                    //                    System.out.println("i = "+i+" j = "+j+" irn transaction ref = "+irn_transaction_reference.get(i)+" fkey ab prod is = "+fkey_abprod.get(j)+" irn transac size is "+
                    //                            irn_transaction_reference.size()+" and conparator 4 size is "+comparator4.size()+" and the current value is and irn realiesed " +
                    //                            "cash flow size is "+irn_realised_cash_flow.size());
                    irn_realised_cash_flow.add(comparator4.get(j));
                    irn_income_value.add(comparator4.get(j));
                    match_found = true;
                }
                if (!match_found && j == fk_realizedcash.size() - 1) {
                    irn_realised_cash_flow.add("0.00");
                    irn_income_value.add("0.00");
                }
            }
        }

        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < irn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (irn_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    irn_market_valuet1.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    irn_market_valuet1.add("0.00");
                }
            }
        }


        //BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE AND UNREALISED SURPLUS VALUE
        BigDecimal out_not_bd, accr_income_bd, marval_bd, appvalue_result, marvalt1_bd, unrealised_result;
        for (int i = 0; i < irn_transaction_reference.size(); i++) {
            try {
                out_not_bd = new BigDecimal(irn_outstanding_notional.get(i));
                accr_income_bd = new BigDecimal(irn_accrued_income.get(i));
                marval_bd = new BigDecimal(irn_market_value.get(i));
                marvalt1_bd = new BigDecimal(irn_market_valuet1.get(i));
                appvalue_result = marval_bd.subtract(out_not_bd).subtract(accr_income_bd);
                irn_appreciation_value.add(String.valueOf(appvalue_result));
                unrealised_result = marval_bd.subtract(marvalt1_bd);
//                irn_unrealised_surplus_value.add(String.valueOf(unrealised_result));
                //UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                                irn_unrealised_surplus_value.add(irn_market_value.get(i));
            } catch (Exception e) {
                //                e.printStackTrace();
                irn_appreciation_value.add("ERROR");
                irn_unrealised_surplus_value.add("ERROR");
            }
        }

        irn_unrealised_surplus_value = negateThatArray(irn_unrealised_surplus_value);
        irn_income_value = negateThatArray(irn_income_value);

        System.out.println("Transaction Reference size is " + irn_transaction_reference.size() + " and the array is " + irn_transaction_reference);
        //        System.out.println("Outstanding Notional size is " + irn_outstanding_notional.size() + " and the array is " + irn_outstanding_notional);
        //        System.out.println("Market Value size is " + irn_market_value.size() + " and the array is " + irn_market_value);
        System.out.println("Book Value size is " + irn_book_value.size() + " and the array is " + irn_book_value);
        System.out.println("Accrued Income Value is " + irn_accrued_income.size() + " and the array is " + irn_accrued_income);
        System.out.println("Realised Cash Flow is " + irn_realised_cash_flow.size() + " and the array is " + irn_realised_cash_flow);
        //        System.out.println("Book Name size is " + irn_bookname.size() + " and the array is " + irn_bookname);
        System.out.println("Appreciation Value size is " + irn_appreciation_value.size() + " and the array is " + irn_appreciation_value);
        System.out.println("Income Value size is " + irn_income_value.size() + " and the array is " + irn_income_value);
        System.out.println("Unrealised Surplus Value size is " + irn_unrealised_surplus_value.size() + " and the array is " + irn_unrealised_surplus_value);
        //        Thread.sleep(2000);
        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for IRN flow - " + timediff(start, end) + " seconds");



    }

    @Test
    public static void IRSFlow() throws Exception {
        //        path = path + date;
        //        tempath = path;
        //        pathminusone = pathminusone + dateminusone;

        product = "IRS";
        instruments.add(product);
        Instant start = Instant.now();
        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        //        filename = "Static Contract Data Report "+date+".csv";
        String filename1 = "Data_Trade_IRS_Bound_" + date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_IRS_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized_Cashflows-") && s.contains(".csv"))
                filename3 = s.substring(s.indexOf("Realized_Cashflows"));
        String f3sheetname = "Realized_Cashflows-" + date;
        String fn_sheetname = "Static Contract Data Report 201";
        String identifier = "IRS";

        ///previous date files
        ArrayList < String > previous_day_files = getFilenamesFromFolder(pathminusone);
        String filename4 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report" + dateminusone + ".csv";
        for (String s: previous_day_files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename4 = s.substring(s.indexOf("AB - PROD"));

        ///// previous date files

        csvToXLSX(path, filename);
        csvToXLSX(path, filename1);
        csvToXLSX(path, filename2);
        csvToXLSX(path, filename3);
        csvToXLSX(pathminusone, filename4);
        //        readSourceFile(path, "Static Contract Data Report 20190329.xlsx", "Static Contract Data Report 201");
        irs_transaction_reference = readPrimaryKey(path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);

        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5, comparator6, comparator7;
        System.out.println(irs_transaction_reference);
        System.out.println("Transaction count size is " + irs_transaction_reference.size() + " and the list is " + irs_transaction_reference);
        //        System.out.println(irs_transaction_reference.toString().replaceAll(",","\n"));

        System.out.println("Reading column 'Pay Outstanding Notional' from file " + path + "\\" + filename1);
        fkey_datatrade = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 24));
        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 25));
        comparator3 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 28));
        comparator4 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 29));
        comparator5 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8));
        comparator6 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7));
        comparator7 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8));
        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3));
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));

        ////// POPULATION OF PON, RON AND PAI, RAI
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (irs_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    irs_pay_outstanding_notional.add(comparator1.get(j));
                    irs_receive_outstanding_notional.add(comparator2.get(j));
                    irs_pay_accrued_income.add(comparator3.get(j));
                    irs_receive_accrued_income.add(comparator4.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    //
                }
            }
        }

        ////// POPULATION OF MARKET VALUE
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod.size(); j++) {
                if (irs_transaction_reference.get(i).equals(fkey_abprod.get(j))) {
                    irs_market_value.add(comparator5.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod.size() - 1) {
                    irs_market_value.add("0.00");
                }
            }
        }


        ////// POPULATION OF YESTERDAY'S MARKET VALUE
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_abprod_yesterday.size(); j++) {
                if (irs_transaction_reference.get(i).equals(fkey_abprod_yesterday.get(j))) {
                    irs_market_valuet1.add(comparator7.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_abprod_yesterday.size() - 1) {
                    irs_market_valuet1.add("0.00");
                }
            }
        }



        //        ////// POPULATION OF REALISED CASH FLOW --BUGGY CODE - DID NOT CONSIDER THE POSSIBILITY OF REALISED_CASHFLOW FILE HAVING MULTIPLE ROWS FOR SAME TRANSACTION
        //        for (int i = 0; i < irs_transaction_reference.size(); i++) {
        //            boolean match_found = false;
        //            for (int j = 0; j < fk_realizedcash.size(); j++) {
        //                if (irs_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
        //                    irs_realised_cash_flow.add(comparator6.get(j));
        //                    match_found = true;
        //                    if (match_found) {
        //                        System.out.println("Found at i = " + i + " and j = " + j + " Conbo is " + irs_transaction_reference.get(i) + "\t" + fk_realizedcash.get(j) + "\t" + comparator6.get(j));
        //                        //                    Thread.sleep(500);
        //                    }
        //                    //                    Thread.sleep(100);
        //                    //                    if(i>100)
        //                    //                        System.exit(1);
        //                    break;
        //                }
        //                if (!match_found && j == fk_realizedcash.size() - 1) {
        //                    irs_realised_cash_flow.add("0.00");
        //                }
        //            }
        //        }


        //TRICKY POPULATION OF REALISED CASH FLOW
        ArrayList < Integer > matchingIndices = new ArrayList < Integer > ();
        ArrayList < String > irs_dupes = new ArrayList < String > ();
        ArrayList < Integer > irs_dupes_indices = new ArrayList < Integer > ();
        ArrayList < Integer > irs_frequency = new ArrayList < Integer > ();
        ArrayList < String > irs_unique_sums = new ArrayList < String > ();
        ArrayList < String > irs_unique_references = new ArrayList < String > ();
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            boolean done = false, match_found = false;
            for (int a = 0; a < fk_realizedcash.size(); a++) {
                String element = fk_realizedcash.get(a);
                if (irs_transaction_reference.get(i).equals(element)) {
                    match_found = true;
                    matchingIndices.add(a);
                    irs_dupes.add(irs_transaction_reference.get(i));
                    //                    System.out.println(Collections.frequency(irs_transaction_reference,irs_transaction_reference.get(i)));
                    int freq = Collections.frequency(fk_realizedcash, irs_transaction_reference.get(i));
                    irs_frequency.add(freq);
                    irs_dupes_indices.add(i);
                    if (freq == 1) {
                        irs_unique_sums.add(comparator6.get(a));
                        irs_unique_references.add(fk_realizedcash.get(a));
                    }
                    if (freq == 2) {
                        if (!irs_unique_references.contains(fk_realizedcash.get(a)))
                            irs_unique_references.add(fk_realizedcash.get(a));
                        BigDecimal v, k;
                        v = new BigDecimal(comparator6.get(a));
                        a++;
                        k = new BigDecimal(comparator6.get(a));
                        irs_unique_sums.add(String.valueOf(v.add(k)));
                        System.out.println("Adding v and k " + v + "\t" + k + " and the result is " + v.add(k));
                        //                        for (int h = 0; h < freq; h++) {
                        //                            if (h == 0)
                        //                                irs_unique_sums.add("puski");
                        //                        }

                    }
                    if (freq == 3) { //this loop is not dev-complete
                        if (!irs_unique_references.contains(fk_realizedcash.get(a)))
                            irs_unique_references.add(fk_realizedcash.get(a));
                        for (int h = 0; h < freq; h++) {
                            if (h == 0)
                                irs_unique_sums.add("puski");
                        }
                    }

                }
            }
            if (match_found) {}
            //                System.out.println("For "+i+" the count is "+matchingIndices);
            //            Thread.sleep(200);
            //            if(matchingIndices.size()>1)
            //                System.exit(1);

            //            if(i>10000)
            //                System.exit(1);
        }

        System.out.println("Matching Indices size " + matchingIndices.size() + " and array is " + matchingIndices);
        System.out.println("IRS Dupes size is " + irs_dupes.size() + " and array is " + irs_dupes);
        System.out.println("IRS Dupes Indices size is " + irs_dupes_indices.size() + " and array is " + irs_dupes_indices);
        System.out.println("IRS Frequency size is " + irs_frequency.size() + " and array is " + irs_frequency);
        System.out.println("IRS unique references size is " + irs_unique_references.size() + " and array is " + irs_unique_references);
        System.out.println("IRS unique sums size is " + irs_unique_sums.size() + " and array is " + irs_unique_sums);

        //TRICKY RE-POPULATION OF IRS INCOME VALUES
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            if (irs_unique_references.contains(irs_transaction_reference.get(i))) {
                int index = irs_unique_references.indexOf(irs_transaction_reference.get(i));
                irs_realised_cash_flow.add(irs_unique_sums.get(index));
            } else
                irs_realised_cash_flow.add("0.00");
        }
        System.out.println("Yipee Kaay Yaay size is " + irs_realised_cash_flow.size() + " and the array is " + irs_realised_cash_flow);





        //BIG DECIMAL CALCULATIONS FOR BOOK VALUE AND ACCRUED INCOME VALUE
        BigDecimal pay_bd, rec_bd, out_notional_result, pai_bd, rai_bd, accrued_income_result;
        for (int i = 0; i < irs_pay_outstanding_notional.size(); i++) {
            try {
                pay_bd = new BigDecimal(irs_pay_outstanding_notional.get(i));
                rec_bd = new BigDecimal(irs_receive_outstanding_notional.get(i));
                pai_bd = new BigDecimal(irs_pay_accrued_income.get(i));
                rai_bd = new BigDecimal(irs_receive_accrued_income.get(i));
                out_notional_result = rec_bd.add(pay_bd);
                irs_book_value.add(String.valueOf(out_notional_result));
                accrued_income_result = rai_bd.add(pai_bd);
                irs_accrued_income_value.add(String.valueOf(accrued_income_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                irs_book_value.add("ERROR");
                irs_accrued_income_value.add("ERROR");
            }
        }

        //BIG DECIMAL CALCULATIONS for APPRECIATION VALUE
        BigDecimal marval_bd, appvalue_result;
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            //            System.out.println("Doing this for values "+irs_pay_outstanding_notional.get(i)+"\n"+irs_receive_outstanding_notional.get(i)+"\n"+irs_pay_accrued_income.get(i)+"\n"+irs_receive_accrued_income.get(i)+"\n"+irs_market_value.get(i));
            try {
                pay_bd = new BigDecimal(irs_pay_outstanding_notional.get(i));
                rec_bd = new BigDecimal(irs_receive_outstanding_notional.get(i));
                pai_bd = new BigDecimal(irs_pay_accrued_income.get(i));
                rai_bd = new BigDecimal(irs_receive_accrued_income.get(i));
                marval_bd = new BigDecimal(irs_market_value.get(i));
                //            System.out.println("And after converting, they are "+pay_bd+"\n"+irs_receive_outstanding_notional.get(i)+"\n"+irs_pay_accrued_income.get(i)+"\n"+irs_receive_accrued_income.get(i)+"\n"+irs_market_value.get(i));
                appvalue_result = marval_bd.subtract(pay_bd).subtract(pai_bd).subtract(rec_bd).subtract(rai_bd);
                irs_appreciation_value.add(String.valueOf(appvalue_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                irs_appreciation_value.add("ERROR");
            }
        }


        //BIG DECIMAL CALCULATION OF UNREALISED SURPLUS VALUE
        BigDecimal marval_today_bd, marval_yesterday_bd, unrealised_result;
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            //            System.out.println("Doing this for values "+irs_pay_outstanding_notional.get(i)+"\n"+irs_receive_outstanding_notional.get(i)+"\n"+irs_pay_accrued_income.get(i)+"\n"+irs_receive_accrued_income.get(i)+"\n"+irs_market_value.get(i));
            try {
                marval_today_bd = new BigDecimal(irs_market_value.get(i));
                marval_yesterday_bd = new BigDecimal(irs_market_valuet1.get(i));
                unrealised_result = marval_today_bd.subtract(marval_yesterday_bd);
//                irs_unrealised_surplus_value.add(String.valueOf(unrealised_result));
                //UPDATING BASED ON VYV;S COMMENTS, FOR A SINGLE DAY, UNREALISED SURPLUS IS = MARKET VALUE FOR THAT SINGLE DAY. IT'S M(t)-M(t-1) ONLY WHEN WE'RE DOING THIS FOR A DATE RANGE
                // UPDATING AGAIN BECAUSE Vyv SAID WHAT HE SAID ON FRIDAY WAS INCORRECT
                                irs_unrealised_surplus_value.add(irs_market_value.get(i));
            } catch (Exception e) {

                //                e.printStackTrace();
                irs_unrealised_surplus_value.add("ERROR");
            }
        }

        irs_unrealised_surplus_value = negateThatArray(irs_unrealised_surplus_value);
        irs_realised_cash_flow = negateThatArray(irs_realised_cash_flow);

        int a = 1;
        try {
            System.out.println("Tried adding some numbers " + Double.parseDouble(irs_pay_outstanding_notional.get(a)) + Double.parseDouble(irs_receive_outstanding_notional.get(a)));
        } catch (Exception e) {
            e.printStackTrace();
        }
        //        System.out.println("Here's a random number "+320000000.00+(-89674587.82541147));
        //        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        //        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        //        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        //        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);
        //        System.out.println("Pay Outstanding Notional size is " + irs_pay_outstanding_notional.size() + " and the array is " + irs_pay_outstanding_notional);
        //        System.out.println("Receive Outstanding Notional size is " + irs_receive_outstanding_notional.size() + " and the array is " + irs_receive_outstanding_notional);
        System.out.println("Book Value size is " + irs_book_value.size() + " and the array is " + irs_book_value);
        System.out.println("Accrued Income Value is " + irs_accrued_income_value.size() + " and the array is " + irs_accrued_income_value);
        //        System.out.println("Market Value is " + irs_market_value.size() + " and the array is " + irs_market_value);
        //        System.out.println("Book Name size is " + irs_book_name.size() + " and the array is " + irs_book_name);
        System.out.println("Appreciation Value size is " + irs_appreciation_value.size() + " and the array is " + irs_appreciation_value);
        System.out.println("Realised Cash Flow size is " + irs_realised_cash_flow.size() + " and the array is " + irs_realised_cash_flow);
        System.out.println("Unrealised Surplus Value size is " + irs_unrealised_surplus_value.size() + " and the array is " + irs_unrealised_surplus_value);
        //        Thread.sleep(2000);
        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for IRS flow - " + timediff(start, end) + " seconds");
    }

    @Test
    public static Object QuintFactTableAPI() throws MalformedURLException, IOException, ParseException {
        System.out.println("Calling Quint Fact Table API");
        HttpURLConnection con = (HttpURLConnection) ((new URL("http://10.122.156.187:14649/EnginesGet/Report?api-key=gXa3l29axz6mghDND/bNtlIpRZuk2r5PaJNMDAZHe2k=&reportname=Complete%20TB%20Report&headers=True&arg1=31%20May%202019&arg2=28%20Jun%202019&Response-Type=csv").openConnection()));
        con.setRequestProperty("Content-Type", "text/csv");
        con.setRequestProperty("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3");
        con.setRequestMethod("GET");
        con.setRequestProperty("Accept-Language", "en-US,en;q=0.9");
        con.setRequestProperty("Accept-Encoding", "gzip,deflate");
        String urlParameters = "api-key=gXa3l29axz6mghDND/bNtlIpRZuk2r5PaJNMDAZHe2k=&reportname=Complete%20TB%20Report&headers=True&arg1=31%20May%202019&arg2=28%20Jun%202019&Response-Type=csv";

        con.setDoOutput(true);
        con.connect();
        DataOutputStream wr = new DataOutputStream(con.getOutputStream());
        wr.writeBytes(urlParameters);
        wr.flush();
        wr.close();
        int responseCode = con.getResponseCode();
        System.out.println("\nSending 'POST' request to URL : ");
//        System.out.println("Post parameters : " + urlParameters);
        System.out.println("Response Code : " + responseCode);
//        BufferedReader in = new BufferedReader(new InputStreamReader(con.getContent()));
        String inputLine;
        StringBuffer response = new StringBuffer();
//        while ((inputLine = in.readLine()) != null) {
//            response.append(inputLine);
//        }
//        in.close();
        // print result
        System.out.println("Quint Fact Table API is: " + response.toString());// 5112018
        return response;

    }


    public static ArrayList < String > getFilenamesFromFolder(String folder_path) {
        //        System.out.println("Begin - getFilenamesFromFolder");
        File folder = new File(folder_path);
        File[] listOfFiles = folder.listFiles();
        ArrayList < String > filenames = new ArrayList < > ();
        System.out.println("Folder path from getFilenamesFromFolder" + folder_path);
        //        System.out.println("list of files  are "+listOfFiles);
        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                //                System.out.println("File " + listOfFiles[i].getName());
                filenames.add(listOfFiles[i].toString());
            } else if (listOfFiles[i].isDirectory()) {
                //                System.out.println("Directory " + listOfFiles[i].getName());
            }
        }
        //        System.out.println("End - getFilenamesFromFolder");
        return filenames;
    }

    public static ArrayList < String > negateThatArray(ArrayList < String > original) {
//        System.out.println("Before negating" + original);
        int minus1 = -1;
        BigDecimal ori, res;
        ArrayList < String > negative = new ArrayList < > ();
        for (int i = 0; i < original.size(); i++) {
            try {
                ori = new BigDecimal(original.get(i));
                res = ori.multiply(new BigDecimal(minus1));
                negative.add(String.valueOf(res));
            } catch (Exception e) {
                //                e.printStackTrace();
                negative.add("NULL");
            }
        }
//        System.out.println("After negating" + negative);
        return negative;
    }

    public static ArrayList < String > getFolderNamesFromDirectory(String directory_path) {
        File folder = new File(directory_path);
        File[] listOfFiles = folder.listFiles();
        ArrayList < String > directorylist = new ArrayList < > ();
        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                //                System.out.println("File " + listOfFiles[i].getName());

            } else if (listOfFiles[i].isDirectory()) {
                directorylist.add(listOfFiles[i].toString());
                //                System.out.println("Directory " + listOfFiles[i].getName());
            }
        }
        //        System.out.println("End - getFilenamesFromFolder");
        return directorylist;
    }

    public static ArrayList < String > readPrimaryKey(String path, String filename, String sheet, String identifier) throws Exception {
        ArrayList < String > pkey = new ArrayList < > ();
        String abs = path + filename;
        setExcelFile(path + "\\" + filename, sheet);
        System.out.println("Reading Primary Key from" + path + " and filename " + filename + " and sheet " + sheet + " - Row count is " + getRowCount(sheet) + " and column count is " + getColumnCount(sheet, 0));
        int headersrow = 3;
        int rcount = getRowCount(sheet);
        int ccount = getColumnCount(sheet, headersrow);
        int i, j, count = 0;
        String celldata;
        ArrayList < String > checker = new ArrayList < > ();
        for (i = 0; i <= rcount; i++) {
            for (j = 0; j < ccount; j++) {
                celldata = getCellData(i, j, sheet);
                checker.add(celldata);
                //                                    System.out.print(celldata+"  |  ");
            }
            if (checker.get(2).startsWith(identifier) && checker.get(2).charAt(3) != 'N') {
                //                System.out.println(i+" | "+checker);
                count++;
                if (!identifier.equals("Cash"))
                    pkey.add(checker.get(2));
                else {
                    if (i > 0 && checker.get(3).equalsIgnoreCase("CRN1245548"))
                        pkey.add(checker.get(3));
                }
                //                    System.out.println("Checking book name thingiri " + daterange_product + " dgfd " + !daterange_product.isEmpty() + checker.get(5));
                try {
                    if (product.equals("IRS"))
                        irs_book_name.add(checker.get(5));
                    if (product.equals("IRN"))
                        irn_bookname.add(checker.get(5));
                    if (product.equals("CRN"))
                        crn_book_name.add(checker.get(5));
                    if (product.equals("ELN"))
                        eln_book_name.add(checker.get(5));
                    if (product.equals("EQF"))
                        eqf_book_name.add(checker.get(5));
                } catch (Exception e) {}
                try {
                    if (daterange_product.equals("IRS"))
                        irs_book_name.add(checker.get(5));
                    if (daterange_product.contains("IRN"))

                        irn_bookname.add(checker.get(5));
                    if (daterange_product.equals("CRN"))
                        crn_book_name.add(checker.get(5));
                    if (daterange_product.equals("ELN(A)"))
                        eln_book_name.add(checker.get(5));
                    if (daterange_product.equals("EQFUT"))
                        eqf_book_name.add(checker.get(5));
                } catch (Exception e) {}


                //                    createOutputFile(checker);
                //                writeColumnData("C:\\Users\\vzk1008\\Documents\\Test Cases\\IRS Test\\","output.xlsx","Transaction Results",0,checker.get(2));
            }
            checker.clear();
            //                if(j==ccount)
            //                    System.out.println();
            //                System.out.println(getCellData(i,0,sheet)+"  |  "+getCellData(i,1,sheet)+"  |  "+getCellData(i,2,sheet)+"  |  "+getCellData(i,3,sheet)+"  |  "+getCellData(i,4,sheet));
        }
        doForSingleTransaction(pkey);
//        System.out.println("Pkey is "+pkey);
        return pkey;
    }

    public static ArrayList<String> doForSingleTransaction(ArrayList<String> in) {

        if(SingleTransactionID.length()>1) {
            in.clear();
        in.add(SingleTransactionID); }
        return in;
    }

    public static ArrayList < String > readColumnData(String path, String filename, String sheet, int colNum) throws Exception {
        ArrayList < String > colData = new ArrayList < > ();
        String out = "null";
        String abs = path + filename;
        setExcelFile(path + "\\" + filename, sheet);
        System.out.println("read column data - Row count is " + getRowCount(sheet) + " and column count is " + getColumnCount(sheet, 0) + "path " + path + "\nfilename " + filename + "\nsheet name" + sheet);
        int headersrow = 3;
        int rcount = getRowCount(sheet);
        int ccount = getColumnCount(sheet, headersrow);
        int i = 0, j, count = 0;
        String celldata;
        ArrayList < String > checker = new ArrayList < > ();
        //        if(product.equalsIgnoreCase("EQF"))
        //            i=5;
        for (; i < rcount; i++) {
            for (j = 0; j <= ccount; j++) { //added = for EQF, need to check if all others are working without an issue
                celldata = getCellData(i, j, sheet);
                checker.add(celldata);
                //                                    System.out.print(celldata+"  |  ");
            }
            //            if (checker.get(0).equals(pkey)) {
            //                                            System.out.println(i+" | "+checker);
            count++;

            //                out = checker.get(colNum);
            colData.add(checker.get(colNum));

            //                    createOutputFile(checker);
            //                writeColumnData("C:\\Users\\vzk1008\\Documents\\Test Cases\\IRS Test\\","output.xlsx","Transaction Results",0,checker.get(2));

            //            }
            checker.clear();
            //                if(j==ccount)
            //                    System.out.println();
            //                System.out.println(getCellData(i,0,sheet)+"  |  "+getCellData(i,1,sheet)+"  |  "+getCellData(i,2,sheet)+"  |  "+getCellData(i,3,sheet)+"  |  "+getCellData(i,4,sheet));
        }

        return colData;
    }


    public static ArrayList<ArrayList < String >> readMultipleColumnData(String path, String filename, String sheet, int colNum1,int colNum2) throws Exception {
        ArrayList<ArrayList<String>> big=new ArrayList<>();
        ArrayList < String > colData1 = new ArrayList < > ();
        ArrayList < String > colData2 = new ArrayList < > ();
        String out = "null";
        String abs = path + filename;
        setExcelFile(path + "\\" + filename, sheet);
        System.out.println("read column data - Row count is " + getRowCount(sheet) + " and column count is " + getColumnCount(sheet, 0) + "path " + path + "\nfilename " + filename + "\nsheet name" + sheet);
        int headersrow = 3;
        int rcount = getRowCount(sheet);
        int ccount = getColumnCount(sheet, headersrow);
        int i = 0, j, count = 0;
        String celldata;
        ArrayList < String > checker = new ArrayList < > ();
        //        if(product.equalsIgnoreCase("EQF"))
        //            i=5;
        for (; i < rcount; i++) {
            for (j = 0; j <= ccount; j++) { //added = for EQF, need to check if all others are working without an issue
                celldata = getCellData(i, j, sheet);
                checker.add(celldata);
                //                                    System.out.print(celldata+"  |  ");
            }
            //            if (checker.get(0).equals(pkey)) {
            //                                            System.out.println(i+" | "+checker);
            count++;

            //                out = checker.get(colNum);
            colData1.add(checker.get(colNum1));
            colData2.add(checker.get(colNum2));

            //                    createOutputFile(checker);
            //                writeColumnData("C:\\Users\\vzk1008\\Documents\\Test Cases\\IRS Test\\","output.xlsx","Transaction Results",0,checker.get(2));

            //            }
            checker.clear();
            //                if(j==ccount)
            //                    System.out.println();
            //                System.out.println(getCellData(i,0,sheet)+"  |  "+getCellData(i,1,sheet)+"  |  "+getCellData(i,2,sheet)+"  |  "+getCellData(i,3,sheet)+"  |  "+getCellData(i,4,sheet));
        }
//        big.add(colData1,colData2);
        return big;
    }

    @BeforeSuite
    public static void archiveTestCases() throws IOException {
        String source_singleday = jenkins_workspace + singleday_jenkins_projectname;
        String source_daterange = jenkins_workspace + daterange_jenkins_projectname;
        String target_singleday = "C:\\Users\\vzk1008\\Documents\\Daily-Run_Source-to-Test-Template_Single-Date Archive\\";
        String target_daterange = "C:\\Users\\vzk1008\\Documents\\Daily-Run_Source-to-Test-Template_Date-Range Archive\\";
        ArrayList < String > list = getFilenamesFromFolder(source_singleday);
        for (String file: list) {
            target_singleday = "C:\\Users\\vzk1008\\Documents\\Daily-Run_Source-to-Test-Template_Single-Date Archive\\";
            target_singleday = file.replace(source_singleday, target_singleday);
            target_daterange = file.replace(source_daterange,target_daterange);

            try {
                Path temp = Files.move(Paths.get(file), Paths.get(target_singleday));
                Path temp1 = Files.move(Paths.get(file),Paths.get(target_daterange));
                if (temp != null) {
                    System.out.println("Moved " + file);
                } else {
                    System.out.println("Move unsuccessful " + file);
                }
            } catch (Exception e) {e.printStackTrace();}

        }

    }

    @AfterSuite
    public static void createOneBigOutputFile() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet;
        XSSFCellStyle bigdecimalstyle = workbook.createCellStyle();
        XSSFDataFormat format = workbook.createDataFormat();
        //            bigdecimalstyle.setDataFormat(format.getFormat("#,##################")); //upto 18 decimal points
        bigdecimalstyle.setDataFormat(format.getFormat("#,##")); //upto 2 decimal points
        //        String fisdle = path + "\\" + product + " Test Template " + date + "_" + Instant.now().toEpochMilli() + ".xlsx";
        String file = jenkins_workspace + singleday_jenkins_projectname + "\\" + "Test Cases for " + date + "_" + Instant.now().toEpochMilli() + ".xlsx";
        System.out.println("starting createOutputFile");
        sheet = workbook.createSheet("Transaction Results");
        int rowNum = 0;
        int colNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(colNum++);
        cell.setCellValue("Transaction Reference");
        cell = row.createCell(colNum++);
        cell.setCellValue("Internal GL Code");
        cell = row.createCell(colNum++);
        cell.setCellValue("Book");
        cell = row.createCell(colNum++);
        //                cell.setCellValue("Pay Outstanding Notional");
        //                cell = row.createCell(colNum++);
        //                cell.setCellValue("Receive Outstanding Notional");
        //                cell = row.createCell(colNum++);
        cell.setCellValue("Book Value");
        cell = row.createCell(colNum++);
        cell.setCellValue("Accrued Income");
        cell = row.createCell(colNum++);
        cell.setCellValue("Appreciation Value");
        cell = row.createCell(colNum++);
        cell.setCellValue("Outstanding Income");
        cell = row.createCell(colNum++);
        cell.setCellValue("Income Value");
        cell = row.createCell(colNum++);
        cell.setCellValue("Unrealised Surplus Value");


        for (String s: instruments) {
            if (s.equals("EQFWD"))
                for (int i = 0; i < eqfwd_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(eqfwd_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(eqfwd_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            if (j == 4) {}
                            if (j == 5) {
                                cell.setCellValue(eqfwd_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            if (j == 7) {
                                cell.setCellValue(eqfwd_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(eqfwd_unrealised_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                        }
                    } catch (Exception e) {}
                }
            if (s.equals("CFS"))
                for (int i = 0; i < cfs_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(cfs_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(cfs_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            if (j == 4) {}
                            if (j == 5) {
                                cell.setCellValue(cfs_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            if (j == 7) {
                                cell.setCellValue(cfs_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(cfs_unrealised_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                        }
                    } catch (Exception e) {}
                }
            if (s.equals("IRFWD"))
                for (int i = 0; i < irfwd_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(irfwd_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(irfwd_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            if (j == 4) {}
                            if (j == 5) {
                                cell.setCellValue(irfwd_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            if (j == 7) {
                                cell.setCellValue(irfwd_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(irfwd_unrealised_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                        }
                    } catch (Exception e) {}
                }
            if (s.equals("CashMI"))
                for (int i = 0; i < cash_transaction_reference_new.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(cash_transaction_reference_new.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(cash_book_name_new.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(cash_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {
                                cell.setCellValue(cash_accrued_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {}
                            if (j == 6) {}
                            if (j == 7) {
                                cell.setCellValue(cash_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {}
                        }
                    } catch (Exception e) {}
                }
            if (s.equals("EQF"))
                for (int i = 0; i < eqf_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(eqf_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(eqf_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 4) {}
                            //                                cell.setCellValue(eqf_income.get(i));
                            if (j == 5) {}
                            //                                cell.setCellValue(irn_appreciation_value.get(i));
                            if (j == 6) {
                                cell.setCellValue(eqf_outstanding_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 7) {
                                cell.setCellValue(eqf_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {}
                        }
                    } catch (Exception e) {}
                }
            if (s.equals("IRN"))
                for (int i = 0; i < irn_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(irn_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(irn_bookname.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(irn_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {
                                cell.setCellValue(irn_accrued_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {
                                cell.setCellValue(irn_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            if (j == 7) {
                                cell.setCellValue(irn_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }
            if (s.equals("ELN"))
                for (int i = 0; i < eln_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(eln_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(eln_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(eln_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {}
                            if (j == 5) {
                                cell.setCellValue(eln_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            if (j == 7) {}
                            if (j == 8) {
                                cell.setCellValue(eln_unreal_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                        }
                    } catch (Exception e) {}
                }
            if (s.equals("CRN"))
                for (int i = 0; i < crn_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(crn_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(crn_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(crn_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {
                                cell.setCellValue(crn_accrued_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {
                                cell.setCellValue(crn_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            if (j == 7) {
                                cell.setCellValue(crn_realised_cash_flow.get(i)); //nothing but income value
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(crn_unreal_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                        }
                    } catch (Exception e) {}
                }
            if (s.equals("IRS"))
                for (int i = 0; i < irs_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(irs_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            if (j == 2) {
                                cell.setCellValue(irs_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(irs_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {
                                cell.setCellValue(irs_accrued_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {
                                cell.setCellValue(irs_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            if (j == 7) {
                                cell.setCellValue(irs_realised_cash_flow.get(i)); //income value
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(irs_unrealised_surplus_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                        }
                    } catch (Exception e) {}
                }
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            workbook.close();


            //            try {
            //                File jen = new File(file);
            //                String jenkinspath = "C:\\Users\\vzk1008\\.jenkins\\workspace\\Daily-Run_Source-to-Test-Template_Single-Date\\";
            //                FileUtils.copyFileToDirectory(jen, new File(jenkinspath));
            //            } catch (Exception e) {}

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    @AfterSuite
    public static void createOneBigDateRangeOutputFile() {
        if(instruments.size()>0) {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet;
            XSSFCellStyle bigdecimalstyle = workbook.createCellStyle();
            XSSFDataFormat format = workbook.createDataFormat();
            //            bigdecimalstyle.setDataFormat(format.getFormat("#,##################")); //upto 18 decimal points
            bigdecimalstyle.setDataFormat(format.getFormat("#,##")); //upto 2 decimal points
            //        String fisdle = path + "\\" + product + " Test Template " + date + "_" + Instant.now().toEpochMilli() + ".xlsx";
            String file = jenkins_workspace + daterange_jenkins_projectname + "\\" + start_date + " - " + end_date + " Test Cases _" + Instant.now().toEpochMilli() + ".xlsx";
            System.out.println("starting createOneBigDateRangeOutputFile for instruments " + instruments);
            sheet = workbook.createSheet("Transaction Results");
            int rowNum = 0;
            int colNum = 0;
            Row row = sheet.createRow(rowNum++);
            Cell cell = row.createCell(colNum++);
            cell.setCellValue("Transaction Reference");
            cell = row.createCell(colNum++);
            cell.setCellValue("Internal GL Code");
            cell = row.createCell(colNum++);
            cell.setCellValue("Book");
            cell = row.createCell(colNum++);
            //                cell.setCellValue("Pay Outstanding Notional");
            //                cell = row.createCell(colNum++);
            //                cell.setCellValue("Receive Outstanding Notional");
            //                cell = row.createCell(colNum++);
            cell.setCellValue("Book Value");
            cell = row.createCell(colNum++);
            cell.setCellValue("Accrued Income");
            cell = row.createCell(colNum++);
            cell.setCellValue("Appreciation Value");
            cell = row.createCell(colNum++);
            cell.setCellValue("Outstanding Income");
            cell = row.createCell(colNum++);
            cell.setCellValue("Income Value");
            cell = row.createCell(colNum++);
            cell.setCellValue("Unrealised Surplus Value");

            for (String s : instruments) {
                if (s.equals("EQFWD"))
                    for (int i = 0; i < eqfwd_daterange_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(eqfwd_daterange_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(eqfwd_book_name.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                }
                                if (j == 4) {
                                }
                                if (j == 5) {
                                    cell.setCellValue(eqfwd_appreciation_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(eqfwd_daterange_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                    cell.setCellValue(eqfwd_daterange_unrealised_surplus.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
                if (s.equals("CFS"))
                    for (int i = 0; i < cfs_daterange_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(cfs_daterange_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(cfs_book_name.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                }
                                if (j == 4) {
                                }
                                if (j == 5) {
                                    cell.setCellValue(cfs_appreciation_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(cfs_daterange_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                    cell.setCellValue(cfs_daterange_unrealised_surplus.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
                if (s.equals("IRFWD"))
                    for (int i = 0; i < irfwd_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(irfwd_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(irfwd_book_name.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                }
                                if (j == 4) {
                                }
                                if (j == 5) {
                                    cell.setCellValue(irfwd_appreciation_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(irfwd_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                    cell.setCellValue(irfwd_unrealised_surplus.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
                if (s.equals("CashMI"))
                    for (int i = 0; i < cash_daterange_transaction_reference_new.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(cash_daterange_transaction_reference_new.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(cash_book_name_new.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                    cell.setCellValue(cash_book_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 4) {
                                    cell.setCellValue(cash_accrued_income.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 5) {
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(cash_daterange_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
                if (s.equals("EQFUT"))
                    for (int i = 0; i < eqf_daterange_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(eqf_daterange_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(eqf_book_name.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                }
                                //                                cell.setCellValue(eqf_outstanding_income.get(i));
                                if (j == 4) {
                                }
                                //                                cell.setCellValue(eqf_income.get(i));
                                if (j == 5) {
                                }
                                //                                cell.setCellValue(irn_appreciation_value.get(i));
                                if (j == 6) {
                                }
                                //                                cell.setCellValue(eqf_outstanding_income.get(i));
                                //                                cell.setCellStyle(bigdecimalstyle);

                                if (j == 7) {
                                    cell.setCellValue(eqf_daterange_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
                if (s.equals("IRN"))
                    for (int i = 0; i < irn_daterange_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(irn_daterange_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(irn_bookname.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                    cell.setCellValue(irn_book_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 4) {
                                    cell.setCellValue(irn_accrued_income.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 5) {
                                    cell.setCellValue(irn_appreciation_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(irn_daterange_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                    cell.setCellValue(irn_unrealised_surplus_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                            }
                        } catch (Exception e) {
                            //                    e.printStackTrace();
                        }
                    }
                if (s.equals("ELN(A)"))
                    for (int i = 0; i < eln_daterange_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(eln_daterange_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(eln_book_name.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                    cell.setCellValue(eln_book_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 4) {
                                }
                                if (j == 5) {
                                    cell.setCellValue(eln_appreciation_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(eln_daterange_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                    cell.setCellValue(eln_daterange_unrealised_surplus.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
                if (s.equals("CRN"))
                    for (int i = 0; i < crn_daterange_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(crn_daterange_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(crn_book_name.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                    cell.setCellValue(crn_book_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 4) {
                                    cell.setCellValue(crn_accrued_income.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 5) {
                                    cell.setCellValue(crn_appreciation_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(crn_daterange_income_value.get(i)); //nothing but income value
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                    cell.setCellValue(crn_daterange_unrealised_surplus.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
                if (s.equals("IRS"))
                    for (int i = 0; i < irs_daterange_transaction_reference.size(); i++) {
                        row = sheet.createRow(rowNum++);
                        colNum = 0;
                        try {
                            for (int j = 0; j <= 10; j++) {
                                cell = row.createCell(colNum++);
                                if (j == 0) {
                                    cell.setCellValue(irs_daterange_transaction_reference.get(i));
                                }
                                if (j == 1) {
                                }
                                if (j == 2) {
                                    cell.setCellValue(irs_book_name.get(i)); //4090312788//15/11/1991//
                                }
                                if (j == 3) {
                                    cell.setCellValue(irs_book_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 4) {
                                    cell.setCellValue(irs_accrued_income_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 5) {
                                    cell.setCellValue(irs_appreciation_value.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 6) {
                                }
                                if (j == 7) {
                                    cell.setCellValue(irs_daterange_income_value.get(i)); //income value
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                                if (j == 8) {
                                    cell.setCellValue(irs_daterange_unrealised_surplus.get(i));
                                    cell.setCellStyle(bigdecimalstyle);
                                }
                            }
                        } catch (Exception e) {
                        }
                    }
            }
            try {
                FileOutputStream outputStream = new FileOutputStream(file);
                workbook.write(outputStream);
                workbook.close();


                //            try {
                //                File jen = new File(file);
                //                String jenkinspath = "C:\\Users\\vzk1008\\.jenkins\\workspace\\Daily-Run_Source-to-Test-Template_Single-Date\\";
                //                FileUtils.copyFileToDirectory(jen, new File(jenkinspath));
                //            } catch (Exception e) {}

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

     public static void createOutputFile() throws IOException {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet;
            XSSFCellStyle bigdecimalstyle = workbook.createCellStyle();
            XSSFDataFormat format = workbook.createDataFormat();
            //            bigdecimalstyle.setDataFormat(format.getFormat("#.##################")); //upto 18 decimal points
            bigdecimalstyle.setDataFormat(format.getFormat("#.##")); //upto 2 decimal points
            //            style.setBorderTop(BorderStyle.DOUBLE);
            //            style.setBorderBottom(BorderStyle.DOUBLE);
            //            style.setFillBackgroundColor(XSSFColor);
            String file = jenkins_workspace + singleday_jenkins_projectname + product + " Test Template " + date + "_" + Instant.now().toEpochMilli() + ".xlsx";
            System.out.println("starting createOutputFile");
            sheet = workbook.createSheet("Transaction Results");
            //            setExcelFile(file, "Transaction Results");
            //            rowNum = getRowCount("Transaction Results");
            int rowNum = 0;
            int colNum = 0;
            ////////////////////////////////////////////////////////////////

            //ADDING THE HEADERS AND DATE RANGE

            Row row = sheet.createRow(rowNum++);
            Cell cell = row.createCell(colNum++);


            if (product.equals("IRS")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");

                // transaction reference, book name, book value, accrued income, appreciation value, outstanding income, income value, unreal surplus
            }

            if (product.equals("CFS")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");

                // transaction reference, book name, book value, accrued income, appreciation value, outstanding income, income value, unreal surplus
            }

            if (product.equals("IRFWD")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");

                // transaction reference, book name, book value, accrued income, appreciation value, outstanding income, income value, unreal surplus
            }


            if (product.equals("IRN")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");
            }


            if (product.equals("CRN")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");
            }

            if (product.equals("ELN")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");
            }


            if (product.equals("EQF")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");
            }

            if (product.equals("CashMI")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");
            }

            if (product.equals("EQFWD")) {
                cell.setCellValue("Transaction Reference");
                cell = row.createCell(colNum++);
                cell.setCellValue("Internal GL Code");
                cell = row.createCell(colNum++);
                cell.setCellValue("Book");
                cell = row.createCell(colNum++);
                //                cell.setCellValue("Pay Outstanding Notional");
                //                cell = row.createCell(colNum++);
                //                cell.setCellValue("Receive Outstanding Notional");
                //                cell = row.createCell(colNum++);
                cell.setCellValue("Book Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Accrued Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Appreciation Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Outstanding Income");
                cell = row.createCell(colNum++);
                cell.setCellValue("Income Value");
                cell = row.createCell(colNum++);
                cell.setCellValue("Unrealised Surplus Value");
            }

            // transaction reference
            // internal GL code
            // book name
            // book value
            // accrued income
            // appreciation value
            // outstanding income
            // income value
            // unreal surplus

            ////////////////////////////////////////////////////////////


            if (product.equals("EQFWD"))
                for (int i = 0; i < eqfwd_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(eqfwd_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(eqfwd_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            //                                cell.setCellValue(cash_book_value.get(i));
                            if (j == 4) {}
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 5) {
                                cell.setCellValue(eqfwd_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 7) {
                                cell.setCellValue(eqfwd_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }

                            if (j == 8) {
                                cell.setCellValue(eqfwd_unrealised_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            //                            if (j == 9){}
                            //                                cell.setCellValue(eqf_income.get(i));
                            //                            if (j == 10) {}
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }

            ////////////////////////////////////////////////////////////


            if (product.equals("CFS"))
                for (int i = 0; i < cfs_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(cfs_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(cfs_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            //                                cell.setCellValue(cash_book_value.get(i));
                            if (j == 4) {}
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 5) {
                                cell.setCellValue(cfs_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 7) {
                                cell.setCellValue(cfs_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(cfs_unrealised_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            //                            if (j == 9){}
                            //                                cell.setCellValue(eqf_income.get(i));
                            //                            if (j == 10) {}
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }


            ////////////////////////////////////////////////////////////

            if (product.equals("IRFWD"))
                for (int i = 0; i < irfwd_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(irfwd_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(irfwd_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            //                                cell.setCellValue(cash_book_value.get(i));
                            if (j == 4) {}
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 5) {
                                cell.setCellValue(irfwd_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 7) {
                                cell.setCellValue(irfwd_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(irfwd_unrealised_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            //                            if (j == 9){}
                            //                                cell.setCellValue(eqf_income.get(i));
                            //                            if (j == 10) {}
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }



            ////////////////////////////////////////////////////////////

            if (product.equals("CashMI"))
                for (int i = 0; i < cash_transaction_reference_new.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(cash_transaction_reference_new.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(cash_book_name_new.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(cash_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {
                                cell.setCellValue(cash_accrued_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {}
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 7) {
                                cell.setCellValue(cash_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            //                            if (j == 9){}
                            //                                cell.setCellValue(eqf_income.get(i));
                            //                            if (j == 10) {}
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }


            ///////////////////////////////////////////////////////////
            if (product.equals("EQF"))
                for (int i = 0; i < eqf_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(eqf_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(eqf_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 4) {}
                            //                                cell.setCellValue(eqf_income.get(i));
                            if (j == 5) {}
                            //                                cell.setCellValue(irn_appreciation_value.get(i));
                            if (j == 6) {
                                cell.setCellValue(eqf_outstanding_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 7) {
                                cell.setCellValue(eqf_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            //                            if (j == 9)
                            //                                cell.setCellValue(eqf_income.get(i));
                            //                            if (j == 10) {}
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }


            //////////////////////////////////////////////////////////////////////////
            if (product.equals("IRN"))
                for (int i = 0; i < irn_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(irn_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(irn_bookname.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(irn_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                                //                                style.setDataFormat(format.getFormat("#,##0,.0000"));
                            }
                            if (j == 4) {
                                cell.setCellValue(irn_accrued_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {
                                cell.setCellValue(irn_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            //                                cell.setCellValue(irn_income_value.get(i));
                            if (j == 7) {
                                cell.setCellValue(irn_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                            if (j == 9)
                            //                                cell.setCellValue(irn_income_value.get(i));
                            //                            if (j == 10)
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }
            ///////////////////////////////////////////////////////////////////

            //////////////////////////////////////////////////////////////////////////
            if (product.equals("ELN"))
                for (int i = 0; i < eln_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(eln_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(eln_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(eln_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {}
                            //                                cell.setCellValue(eln_appreciation_value.get(i));
                            if (j == 5) {
                                cell.setCellValue(eln_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            //                                cell.setCellValue(irn_income_value.get(i));
                            if (j == 7) {}
                            //                                cell.setCellValue(eln_unreal_surplus.get(i));
                            if (j == 8) {
                                cell.setCellValue(eln_unreal_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                            if (j == 9)
                            //                                cell.setCellValue(irn_income_value.get(i));
                            //                            if (j == 10)
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }
            ///////////////////////////////////////////////////////////////////


            //////////////////////////////////////////////////////////////////////////
            if (product.equals("CRN"))
                for (int i = 0, l = 2; i < crn_transaction_reference.size(); i++, l++) {
                    row = sheet.createRow(rowNum++);
                    String strFormula = "-=SUMIFS('Realized_Cashflows-20190605.xlsx'!$H:$H;'Realized_Cashflows-20190605.xlsx'!$D:$D;A" + l + ")";
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(crn_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(crn_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(crn_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {
                                cell.setCellValue(crn_accrued_income.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {
                                cell.setCellValue(crn_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            //                                cell.setCellValue(crn_realised_cash_flow.get(i));
                            if (j == 7) {
                                cell.setCellValue(crn_realised_cash_flow.get(i)); //nothing but income value
                                //                                cell.setCellFormula(strFormula);
                                cell.setCellStyle(bigdecimalstyle);
                                //                                System.out.println(crn_transaction_reference.get(i)+"\t"+crn_realised_cash_flow.get(i));
                            }
                            if (j == 8) {
                                cell.setCellValue(crn_unreal_surplus.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                            if (j == 9)
                            //                                cell.setCellValue(irn_income_value.get(i));
                            //                            if (j == 10)
                            //                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }
            ///////////////////////////////////////////////////////////////////


            /////////////////////////////////////////////////////////////////
            if (product.equals("IRS"))
                for (int i = 0; i < irs_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0) {
                                cell.setCellValue(irs_transaction_reference.get(i));
                            }
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2) {
                                cell.setCellValue(irs_book_name.get(i)); //4090312788//15/11/1991//
                            }
                            if (j == 3) {
                                cell.setCellValue(irs_book_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 4) {
                                cell.setCellValue(irs_accrued_income_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 5) {
                                cell.setCellValue(irs_appreciation_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 6) {}
                            //                                cell.setCellValue(irs_realised_cash_flow.get(i));
                            if (j == 7) {
                                cell.setCellValue(irs_realised_cash_flow.get(i)); //income value
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            if (j == 8) {
                                cell.setCellValue(irs_unrealised_surplus_value.get(i));
                                cell.setCellStyle(bigdecimalstyle);
                            }
                            //                            if (j == 9)
                            //                                cell.setCellValue(irs_realised_cash_flow.get(i));
                            //                            if (j == 10)
                            //                                cell.setCellValue(irs_unrealised_surplus_value.get(i));
                        }
                    } catch (Exception e) {
                        //                    e.printStackTrace();
                    }
                }




            try {
                FileOutputStream outputStream = new FileOutputStream(file);
                workbook.write(outputStream);
                workbook.close();


                try {
                    File jen = new File(file);
                    String directorypath = "C:\\Users\\vzk1008\\.jenkins\\workspace\\Daily-Run_Source-to-Test-Template-Files\\";
                    FileUtils.copyFileToDirectory(jen, new File(directorypath));
                } catch (Exception e) {}

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //        catch (FileNotFoundException e) {
        //            e.printStackTrace();
        //        } catch (IOException e) {
        //            e.printStackTrace();
        //        }
        catch (Exception e) {
            e.printStackTrace();
        }

    }

}
