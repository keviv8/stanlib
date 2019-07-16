package liberty;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.*;

import java.io.*;
import java.math.BigDecimal;
import java.time.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

import static liberty.ExcelUtil.*;


public class AutoTest {
    public static Date realdate = new Date();
    public static LocalDate localDate = realdate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    public static int year = localDate.getYear();
    public static Month month = localDate.getMonth();
    public static int day = localDate.getDayOfMonth();
    public static boolean opfilecreated = false;
    public static String path;
    public static String tempath;
    public static String jdrive = "\\\\libfin01\\libfin\\Libfin\\Systems\\UniCalc\\Production\\A2 Archive\\" + year + "\\" + month.toString() + "\\";
    public static String pathminusone;
    public static String tempathminusone;
    public static String date;
    public static String dateminusone;
    public static String product = null;
    public static ArrayList < String > daysgone = new ArrayList < > ();

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
            daysgone.add(LocalDate.now().minusDays(i).toString().replace("-",""));
        }
        System.out.println(daysgone);
        System.out.println(getDateRanges("20190531","20190404"));
//        System.out.println(months);


        //        csvToXLSX(path,"Static Contract Data Report 20190329.csv");
        //        System.out.println("CSV conversion done, waiting for a couple seconds");
        //        ArrayList<String> files = getFilenamesFromFolder(path);
        //        System.out.println(files);
        //        Thread.sleep(9000);

        //        readSourceFile(path, filename, sheet);

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

    public static List<String> getDateRanges(String from, String to) {
        ArrayList <String> dr = new ArrayList<>();
        int from_index,to_index;
        for(String s:daysgone) {
            if(s.equalsIgnoreCase(from))
                from_index=daysgone.indexOf(from);
        }

        return daysgone.subList(daysgone.indexOf(from),daysgone.indexOf(to)+1);
    }

    @BeforeMethod
    public static void beforeMethod() {
        path = "C:\\Users\\vzk1008\\Documents\\04 production\\";
        date = "20190604";
        pathminusone = path;
        dateminusone = "20190603";
        path = path + date;
        tempath = path;
        pathminusone = pathminusone + dateminusone;
    }

    @BeforeSuite
    public static void calculateDays() {
        for (int i = 0; i < 1000; i++) {
            daysgone.add(LocalDate.now().minusDays(i).toString().replace("-",""));
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
            if (s.contains("Realized Cashflows-") && s.contains(".csv"))
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
            if (s.contains("Realized Cashflows-") && s.contains(".csv"))
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
        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
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
                    eqf_outstanding_income.add("NULL");
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
                eqf_income.add("NULL");
            }
        }


        System.out.println("Transaction reference  size is " + eqf_transaction_reference.size() + " and the array is " + eqf_transaction_reference);
        System.out.println("Book Name sizeis " + eqf_book_name.size() + " and the array is " + eqf_book_name);
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
            if (s.contains("Realized Cashflows-") && s.contains(".csv"))
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
        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        for (int i = 2; i < comparator5.size(); i++) {
            cash_transaction_reference.add(comparator5.get(i));
            cash_book_name.add(comparator4.get(i));
            cash_book_value.add(comparator2.get(i));
            cash_accrued_income.add(comparator3.get(i));
            cash_income.add(comparator1.get(i));
        }

        for(String s : cash_transaction_reference) {
            String n = s.replace("\"","");
            cash_transaction_reference_new.add(n);
        }

        for(String s : cash_book_name) {
            String n = s.replace("\"","");
            cash_book_name_new.add(n);
        }

        negateThatArray(cash_income);

        System.out.println("Transaction Reference size is " + cash_transaction_reference.size() + " and the array is " + cash_transaction_reference);
        System.out.println("Book Name size is " + cash_book_name.size() + " and the array is " + cash_book_name);
        System.out.println("Accrued Income size is " + cash_accrued_income.size() + " and the array is " + cash_accrued_income);
        System.out.println("Book Value size is " + cash_book_value.size() + " and the array is " + cash_book_value);
        System.out.println("Income size is " + cash_income.size() + " and the array is " + cash_income);




        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for CashMI flow - " + timediff(start, end) + " seconds");

    }


    @Test
    public static void ELNFlow() throws Exception {
        //        path = path + date;
        //        tempath = path;
        //        pathminusone = pathminusone + dateminusone;
        product = "ELN";
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
            if (s.contains("Realized Cashflows-") && s.contains(".csv"))
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
        System.out.println("Transaction count size is " + eln_transaction_reference.size() + " and the list is " + eln_transaction_reference);
        fkey_datatrade = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 0));
        comparator1 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 7)); //outstanding notional or bookvalue
        //        comparator2 = (readColumnData(path, filename1.replace(".csv", ".xlsx"), filename1.replace(".csv", ""), 17)); // accrued intetest native
        comparator3 = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 8)); //market value column//i column in ab prod
        comparator4 = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 7)); //
        comparator5 = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 8)); //yesterday's market value
        //        fk_realizedcash = (readColumnData(path, filename3.replace(".csv", ".xlsx"), filename3.replace(".csv", ""), 3)); //
        fkey_abprod = (readColumnData(path, filename2.replace(".csv", ".xlsx"), f2sheetname, 5));
        fkey_abprod_yesterday = (readColumnData(pathminusone, filename4.replace(".csv", ".xlsx"), f2sheetname, 5));

        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        ////// POPULATION OF BOOK VALUE
        for (int i = 0; i < eln_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (eln_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    eln_book_value.add(comparator1.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    eln_book_value.add("NULL");
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
                    eln_market_value.add("NULL");
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
                    eln_market_valuet1.add("NULL");
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
                eln_unreal_surplus.add(String.valueOf(unrealised_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                eln_unreal_surplus.add("ERROR");
            }
        }


        negateThatArray(eln_unreal_surplus);

        System.out.println("Transaction Reference size is " + eln_transaction_reference.size() + " and the array is " + eln_transaction_reference);
        System.out.println("Book Name size is " + eln_book_name.size() + " and the array is " + eln_book_name);
        System.out.println("Market Value size is " + eln_market_value.size() + " and the array is " + eln_market_value);
        System.out.println("Book Value size is " + eln_book_value.size() + " and the array is " + eln_book_value);
        System.out.println("Yesterday's Market Value size is " + eln_market_valuet1.size() + " and the array is " + eln_market_valuet1);
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
        Instant start = Instant.now();

        String filename = "Static Contract Data Report 20190402.csv";
        ArrayList < String > files = getFilenamesFromFolder(path);
        for (String s: files)
            if (s.contains("Static Contract Data Report") && s.contains(".csv"))
                filename = s.substring(s.indexOf("Static"));
        String filename1 = "Data_Trade_CRN_Bound_" + date + ".csv";
        for (String s: files)
            if (s.contains("Data_Trade_CRN_Bound_2019") && s.contains(".csv"))
                filename1 = s.substring(s.indexOf("Data_Trade"));
        String filename2 = "AB - PROD - FIN - BO Libfin Native Currencies Valuations Report " + date + ".csv";
        for (String s: files)
            if (s.contains("AB - PROD - FIN - BO Libfin Native Currencies Valuations Report") && s.contains(".csv"))
                filename2 = s.substring(s.indexOf("AB - PROD"));

        String f2sheetname = "AB - PROD - FIN - BO Libfin Nat";
        String filename3 = "Realized_Cashflows-" + date + ".csv";
        for (String s: files)
            if (s.contains("Realized Cashflows-") && s.contains(".csv"))
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

        //        csvToXLSX(path, filename);
        csvToXLSX(path, filename1);
        //        csvToXLSX(path, filename2);
        //        csvToXLSX(path, filename3);
        //        csvToXLSX(pathminusone, filename4);

        crn_transaction_reference = readPrimaryKey(path, filename.replace(".csv", ".xlsx"), fn_sheetname, identifier);
        ArrayList < String > fkey_datatrade, fkey_abprod, fkey_abprod_yesterday, fk_realizedcash, comparator1, comparator2, comparator3, comparator4, comparator5;
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

        //        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        //        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        //        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        //        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        //        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        //        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        //        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


        ////// POPULATION OF BOOK VALUE AND ACCRUED INCOME
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fkey_datatrade.size(); j++) {
                if (crn_transaction_reference.get(i).equals(fkey_datatrade.get(j))) {
                    crn_book_value.add(comparator1.get(j));
                    crn_accrued_income.add(comparator2.get(j));
                    match_found = true;
                }
                if (!match_found && j == fkey_datatrade.size() - 1) {
                    crn_book_value.add("NULL");
                    crn_accrued_income.add("NULL");
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
                    crn_market_value.add("NULL");
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
                    crn_market_valuet1.add("NULL");
                }
            }
        }

        ////// POPULATION OF REALISED CASH FLOW
        for (int i = 0; i < crn_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fk_realizedcash.size(); j++) {
                if (crn_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
                    crn_realised_cash_flow.add(comparator4.get(j));
                    match_found = true;
                }
                if (!match_found && j == fk_realizedcash.size() - 1) {
                    crn_realised_cash_flow.add("NULL");
                }
            }
        }

        //// BIG DECIMAL CALCULATIONS FOR APPRECIATION VALUE
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
                crn_unreal_surplus.add(String.valueOf(unrealised_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                crn_appreciation_value.add("ERROR");
                crn_unreal_surplus.add("ERROR");
            }
        }

        negateThatArray(crn_unreal_surplus);
        negateThatArray(crn_realised_cash_flow);

        System.out.println("Transaction Reference size is " + crn_transaction_reference.size() + " and the array is " + crn_transaction_reference);
        System.out.println("Market Value size is " + crn_market_value.size() + " and the array is " + crn_market_value);
        System.out.println("Book Value size is " + crn_book_value.size() + " and the array is " + crn_book_value);
        System.out.println("Accrued Income Value is " + crn_accrued_income.size() + " and the array is " + crn_accrued_income);
        System.out.println("Book Name size is " + crn_book_name.size() + " and the array is " + crn_book_name);
        System.out.println("Appreciation Value size is " + crn_appreciation_value.size() + " and the array is " + crn_appreciation_value);
        System.out.println("Yesterday's Market Value size is " + crn_market_valuet1.size() + " and the array is " + crn_market_valuet1);
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
            if (s.contains("Realized Cashflows-") && s.contains(".csv"))
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


        System.out.println("Fkey DT size is " + fkey_datatrade.size() + " and the list is " + fkey_datatrade);
        System.out.println("Fkey AB PROD size is " + fkey_abprod.size() + " and the list is " + fkey_abprod);
        System.out.println("Comparator 1 size is " + comparator1.size() + " and the list is " + comparator1);
        System.out.println("Comparator 2 size is " + comparator2.size() + " and the list is " + comparator2);
        System.out.println("Comparator 3 size is " + comparator3.size() + " and the list is " + comparator3);
        System.out.println("Comparator 4 size is " + comparator4.size() + " and the list is " + comparator4);
        System.out.println("Comparator 5 size is " + comparator5.size() + " and the list is " + comparator5);


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
                    irn_outstanding_notional.add("NULL");
                    irn_book_value.add("NULL");
                    irn_accrued_income.add("NULL");
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
                    irn_market_value.add("NULL");
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
                    irn_realised_cash_flow.add("NULL");
                    irn_income_value.add("NULL");
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
                    irn_market_valuet1.add("NULL");
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
                irn_unrealised_surplus_value.add(String.valueOf(unrealised_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                irn_appreciation_value.add("ERROR");
                irn_unrealised_surplus_value.add("ERROR");
            }
        }

        negateThatArray(irn_unrealised_surplus_value);
        negateThatArray(irn_income_value);

        System.out.println("Transaction Reference size is " + irn_transaction_reference.size() + " and the array is " + irn_transaction_reference);
        System.out.println("Outstanding Notional size is " + irn_outstanding_notional.size() + " and the array is " + irn_outstanding_notional);
        System.out.println("Market Value size is " + irn_market_value.size() + " and the array is " + irn_market_value);
        System.out.println("Book Value size is " + irn_book_value.size() + " and the array is " + irn_book_value);
        System.out.println("Accrued Income Value is " + irn_accrued_income.size() + " and the array is " + irn_accrued_income);
        System.out.println("Realised Cash Flow is " + irn_realised_cash_flow.size() + " and the array is " + irn_realised_cash_flow);
        System.out.println("Book Name size is " + irn_bookname.size() + " and the array is " + irn_bookname);
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
            if (s.contains("Realized Cashflows-") && s.contains(".csv"))
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
                    irs_market_value.add("NULL");
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
                    irs_market_valuet1.add("NULL");
                }
            }
        }



        ////// POPULATION OF REALISED CASH FLOW
        for (int i = 0; i < irs_transaction_reference.size(); i++) {
            boolean match_found = false;
            for (int j = 0; j < fk_realizedcash.size(); j++) {
                if (irs_transaction_reference.get(i).equals(fk_realizedcash.get(j))) {
                    irs_realised_cash_flow.add(comparator6.get(j));
                    match_found = true;
                }
                if (!match_found && j == fk_realizedcash.size() - 1) {
                    irs_realised_cash_flow.add("NULL");
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
                irs_unrealised_surplus_value.add(String.valueOf(unrealised_result));
            } catch (Exception e) {
                //                e.printStackTrace();
                irs_unrealised_surplus_value.add("ERROR");
            }
        }

        negateThatArray(irs_unrealised_surplus_value);
        negateThatArray(irs_realised_cash_flow);

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
        System.out.println("Pay Outstanding Notional size is " + irs_pay_outstanding_notional.size() + " and the array is " + irs_pay_outstanding_notional);
        System.out.println("Receive Outstanding Notional size is " + irs_receive_outstanding_notional.size() + " and the array is " + irs_receive_outstanding_notional);
        System.out.println("Book Value size is " + irs_book_value.size() + " and the array is " + irs_book_value);
        System.out.println("Accrued Income Value is " + irs_accrued_income_value.size() + " and the array is " + irs_accrued_income_value);
        System.out.println("Market Value is " + irs_market_value.size() + " and the array is " + irs_market_value);
        System.out.println("Book Name size is " + irs_book_name.size() + " and the array is " + irs_book_name);
        System.out.println("Appreciation Value size is " + irs_appreciation_value.size() + " and the array is " + irs_appreciation_value);
        System.out.println("Realised Cash Flow size is " + irs_realised_cash_flow.size() + " and the array is " + irs_realised_cash_flow);
        System.out.println("Unrealised Surplus Value size is " + irs_unrealised_surplus_value.size() + " and the array is " + irs_unrealised_surplus_value);
        Thread.sleep(2000);
        createOutputFile();
        Instant end = Instant.now();
        System.out.println("Time taken for IRS flow - " + timediff(start, end) + " seconds");
    }


    public static String saveOutputToLocal(File filename) {
        try {
            Date date = new Date();
            LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            int year = localDate.getYear();
            Month month = localDate.getMonth();
            int day = localDate.getDayOfMonth();
            String directorypath = "C:\\Users\\vkolla\\Documents\\Processed Source Files\\" + year + "\\" + month + "\\" + day + "\\";
            new File(directorypath).mkdirs();
            String path = directorypath + filename;
            FileUtils.copyFileToDirectory(filename, new File(directorypath));
        } catch (Exception e) {

        }
        return path;
    }


    public static ArrayList < String > getFilenamesFromFolder(String folder_path) {
        //        System.out.println("Begin - getFilenamesFromFolder");
        File folder = new File(folder_path);
        File[] listOfFiles = folder.listFiles();
        ArrayList < String > filenames = new ArrayList < > ();
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

    public static ArrayList<String > negateThatArray(ArrayList<String> original) {
        int minus1 = -1;
        BigDecimal ori,res;
        ArrayList<String > negative = new ArrayList<>();
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
        System.out.println("Row count is " + getRowCount(sheet) + " and column count is " + getColumnCount(sheet, 0));
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
                    if (i > 0)
                        pkey.add(checker.get(3));
                }
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
                //                    createOutputFile(checker);
                //                writeColumnData("C:\\Users\\vzk1008\\Documents\\Test Cases\\IRS Test\\","output.xlsx","Transaction Results",0,checker.get(2));
            }
            checker.clear();
            //                if(j==ccount)
            //                    System.out.println();
            //                System.out.println(getCellData(i,0,sheet)+"  |  "+getCellData(i,1,sheet)+"  |  "+getCellData(i,2,sheet)+"  |  "+getCellData(i,3,sheet)+"  |  "+getCellData(i,4,sheet));
        }
        return pkey;
    }

    public static ArrayList < String > readColumnData(String path, String filename, String sheet, int colNum) throws Exception {
        ArrayList < String > colData = new ArrayList < > ();
        String out = "null";
        String abs = path + filename;
        setExcelFile(path + "\\" + filename, sheet);
        System.out.println("read column data - Row count is " + getRowCount(sheet) + " and column count is " + getColumnCount(sheet, 0));
        //        System.out.println("path "+path+"\nfilename "+filename+"\nsheet name"+sheet);
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

    public static void readSourceFile(String path, String filename, String sheet) {
        int headrow = 0;
        try {
            setExcelFile(path + filename, sheet);
            System.out.println("Row count is " + getRowCount(sheet) + " and column count is " + getColumnCount(sheet, 0));
            int rcount = getRowCount(sheet);
            int ccount = getColumnCount(sheet, headrow);
            int i, j, count = 0;
            String search_parameter = "IRS";
            String celldata;
            ArrayList < String > checker = new ArrayList < > ();
            for (i = 0; i < rcount; i++) {
                for (j = 0; j < ccount; j++) {
                    celldata = getCellData(i, j, sheet);
                    checker.add(celldata);
                    //                    System.out.print(celldata+"  |  ");
                }
                if (checker.get(2).startsWith(search_parameter)) {
                    System.out.println(i + " | " + checker);
                    count++;
                    //                    createOutputFile(checker);
                    writeColumnData("C:\\Users\\vzk1008\\Documents\\Test Cases\\IRS Test\\", "output.xlsx", "Transaction Results", 0, checker.get(2));

                }
                checker.clear();
                //                if(j==ccount)
                //                    System.out.println();
                //                System.out.println(getCellData(i,0,sheet)+"  |  "+getCellData(i,1,sheet)+"  |  "+getCellData(i,2,sheet)+"  |  "+getCellData(i,3,sheet)+"  |  "+getCellData(i,4,sheet));
            }
            System.out.println("Found " + count + " records in the file \"" + filename + "\"");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    public static void writeColumnData(String filepath, String filename, String sheet, int colNum, String cellvalue) throws Exception {
        setExcelFile(filepath + filename, sheet);
        setCellData(cellvalue, getRowCount("Transaction Results"), colNum, "Transaction Results", filepath + filename);

    }


    public static void createOutputFile() throws IOException {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();;
            XSSFSheet sheet;
            XSSFCellStyle bigdecimalstyle = workbook.createCellStyle();
//            bigdecimalstyle.setDataFormat(cr);
            XSSFDataFormat format = workbook.createDataFormat();
//            style.setBorderTop(BorderStyle.DOUBLE);
//            style.setBorderBottom(BorderStyle.DOUBLE);
            //            style.setFillBackgroundColor(XSSFColor);
            String file = path + "\\" + product + " Test Template " + date + "_" + Instant.now().toEpochMilli() + ".xlsx";
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


            if (product.equals("CashMI"))
                for (int i = 0; i < cash_transaction_reference_new.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0)
                                cell.setCellValue(cash_transaction_reference_new.get(i));
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2)
                                cell.setCellValue(cash_book_name_new.get(i)); //4090312788//15/11/1991//
                            if (j == 3)
                                cell.setCellValue(cash_book_value.get(i));
                            if (j == 4)
                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 5) {}
                            //                                cell.setCellValue(cash_accrued_income.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 7)
                                cell.setCellValue(cash_income.get(i));
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
                            if (j == 0)
                                cell.setCellValue(eqf_transaction_reference.get(i));
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2)
                                cell.setCellValue(eqf_book_name.get(i)); //4090312788//15/11/1991//
                            if (j == 3) {}
                            //                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 4) {}
                            //                                cell.setCellValue(eqf_income.get(i));
                            if (j == 5) {}
                            //                                cell.setCellValue(irn_appreciation_value.get(i));
                            if (j == 6)
                                cell.setCellValue(eqf_outstanding_income.get(i));
                            if (j == 7)
                                cell.setCellValue(eqf_income.get(i));
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
                            if (j == 0)
                                cell.setCellValue(irn_transaction_reference.get(i));
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2)
                                cell.setCellValue(irn_bookname.get(i)); //4090312788//15/11/1991//
                            if (j == 3) {
                                cell.setCellValue(irn_book_value.get(i));

//                                style.setDataFormat(format.getFormat("#,##0,.0000"));
                            }
                            if (j == 4)
                                cell.setCellValue(irn_accrued_income.get(i));
                            if (j == 5)
                                cell.setCellValue(irn_appreciation_value.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(irn_income_value.get(i));
                            if (j == 7)
                                cell.setCellValue(irn_income_value.get(i));
                            if (j == 8)
                                cell.setCellValue(irn_unrealised_surplus_value.get(i));
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
                            if (j == 0)
                                cell.setCellValue(eln_transaction_reference.get(i));
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2)
                                cell.setCellValue(eln_book_name.get(i)); //4090312788//15/11/1991//
                            if (j == 3)
                                cell.setCellValue(eln_book_value.get(i));
                            if (j == 4) {}
                            //                                cell.setCellValue(eln_appreciation_value.get(i));
                            if (j == 5)
                                cell.setCellValue(eln_appreciation_value.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(irn_income_value.get(i));
                            if (j == 7) {}
                            //                                cell.setCellValue(eln_unreal_surplus.get(i));
                            if (j == 8)
                                cell.setCellValue(eln_unreal_surplus.get(i));
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
                for (int i = 0; i < crn_transaction_reference.size(); i++) {
                    row = sheet.createRow(rowNum++);
                    colNum = 0;
                    try {
                        for (int j = 0; j <= 10; j++) {
                            cell = row.createCell(colNum++);
                            if (j == 0)
                                cell.setCellValue(crn_transaction_reference.get(i));
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2)
                                cell.setCellValue(crn_book_name.get(i)); //4090312788//15/11/1991//
                            if (j == 3)
                                cell.setCellValue(crn_book_value.get(i));
                            if (j == 4)
                                cell.setCellValue(crn_accrued_income.get(i));
                            if (j == 5)
                                cell.setCellValue(crn_appreciation_value.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(crn_realised_cash_flow.get(i));
                            if (j == 7)
                                cell.setCellValue(crn_realised_cash_flow.get(i)); //nothing but income value
                            if (j == 8)
                                cell.setCellValue(crn_unreal_surplus.get(i));
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
                            if (j == 0)
                                cell.setCellValue(irs_transaction_reference.get(i));
                            if (j == 1) {}
                            //YET TO IMPLEMENT INTERNAL GL CODE
                            if (j == 2)
                                cell.setCellValue(irs_book_name.get(i)); //4090312788//15/11/1991//
                            if (j == 3)
                                cell.setCellValue(irs_book_value.get(i));
                            if (j == 4)
                                cell.setCellValue(irs_accrued_income_value.get(i));
                            if (j == 5)
                                cell.setCellValue(irs_appreciation_value.get(i));
                            if (j == 6) {}
                            //                                cell.setCellValue(irs_realised_cash_flow.get(i));
                            if (j == 7)
                                cell.setCellValue(irs_realised_cash_flow.get(i)); //income value
                            if (j == 8)
                                cell.setCellValue(irs_unrealised_surplus_value.get(i));
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