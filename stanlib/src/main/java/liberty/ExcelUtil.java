package liberty;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.DateFormatConverter;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.*;


public class ExcelUtil {
    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    static XSSFCell Cell;
    private static XSSFRow Row;
    public static XSSFWorkbook wb;
    public static XSSFSheet sheet;
    public static XSSFRow row;
    public static XSSFCell cell;
    public static FileInputStream fis;

    @SuppressWarnings("static-access")
    // This method is to set the File path and to open the Excel file
    // Pass Excel Path and SheetName as Arguments to this method
    public static void setExcelFile(String Path, String SheetName) throws Exception {
        FileInputStream ExcelFile = new FileInputStream(Path);
        ExcelWBook = new XSSFWorkbook(ExcelFile);
        ExcelWSheet = ExcelWBook.getSheet(SheetName);
    }


    // This method is to read the test data from the Excel cell
    // In this we are passing Arguments as Row Num, Col Num & Sheet Name
    @SuppressWarnings("static-access")
    public static String getCellData(int RowNum, int ColNum, String SheetName) throws Exception {
        ExcelWSheet = ExcelWBook.getSheet(SheetName);
        String CellData;
        try {
            Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
            if (Cell.getCellType() == CellType.STRING) {
                Cell.setCellType(CellType.STRING);
            }

            CellData = Cell.getStringCellValue();
            //            if(HSSFDateUtil.isCellDateFormatted(Cell)) {
            //                Date celldata = Cell.getDateCellValue();
            //                DateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd hh:mm:ss");
            //                CellData = dateFormat.format(celldata);
            //            }
            //            System.out.println(getCellStyle(RowNum, ColNum, SheetName));
            return CellData;
        } catch (NullPointerException e) {
            return "null";
        }
    }

    public static String csvLineCleanUp(String currentLine) {
        //        System.out.println("Line with commas\t\t"+currentLine);
        String answer = null;
        if(!currentLine.contains("\",\"")) {
        int firstquote = currentLine.indexOf("\"");
        String temp = currentLine.substring(firstquote + 1);
        //        System.out.println("Temp string\t\t"+temp);
        int nextquote = temp.indexOf("\"");
        String badcell = temp.substring(0, nextquote);
        //        System.out.println("Bad cell\t\t"+badcell);
        String goodcell = badcell.replaceAll(",", ":::");

        //        System.out.println("Bad cell updated\t\t"+goodcell);
        String goodLine = currentLine.replace(badcell, goodcell);
        answer= goodLine;
        }
        else
            answer=currentLine;

        //        System.exit(1);
        return answer;
    }

    public static ArrayList< String > getFilenamesFromFolder(String folder_path) {
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

    public static boolean fileAlreadyExists(String path,String filename) {
        boolean result=false;
        String full = path+"\\"+filename;
        full = full.replace("\"\"","\"");
        ArrayList <String> files  = getFilenamesFromFolder(path);
        for(String s:files) {
//            int lastslashins = s.lastIndexOf("\"");
//            String justfilenameins=s.substring(lastslashins);
//            String lastslashinfull = full.
//            System.out.println("Comparing '"+full+"' with '"+s+"'"+result);
            if(s.equalsIgnoreCase(full)) {
                result = true;

            }
            if(s.contains(filename)){
                result = true;
//                        System.out.println("File's already there");
                }
        }

        return result;
    }

    public static float timediff(Instant from, Instant to) {
        return (float)((float)(Duration.between(from, to).toMillis()) / (float) 1000);
    }

    public static File csvToXLSX(String path, String filename) {
        File myFile=null;
        try {
            Instant start = Instant.now();
            String xlsxFileAddress = null;
            ArrayList<String> filelist = getFilenamesFromFolder(path);
            String temp = filename.replace(".csv", ".xlsx");
//            System.out.println(filename+"checking if this already exists "+temp+"yoo "+fileAlreadyExists(path,temp));

            if(!fileAlreadyExists(path,temp)) {
                System.out.print("Converting file " + filename + " at location " + path);

                String csvFileAddress = path + "\\" + filename.replace(".xlsx", ".csv"); //csv file address
                xlsxFileAddress = path + "\\" + filename.replace(".csv", ".xlsx"); //xlsx file address
                XSSFWorkbook workBook = new XSSFWorkbook();
                XSSFSheet sheet = workBook.createSheet(filename.replace(".csv", ""));
                String currentLine = null;
                int RowNum = 0;
                BufferedReader br = new BufferedReader(new FileReader(csvFileAddress));
                while ((currentLine = br.readLine()) != null) {
                    if (currentLine.contains("\"")) {
                        //                    System.out.println("Before line cleanup\t\t\t"+currentLine);
                        currentLine = csvLineCleanUp(currentLine);
                        //                    System.out.println("Aftere line cleanup\t\t\t"+currentLine);
                        //                    System.exit(0);
                    }
                    String str[] = currentLine.split(",");
                    RowNum++;
                    XSSFRow currentRow = sheet.createRow(RowNum);
                    for (int i = 0; i < str.length; i++) {
                        currentRow.createCell(i).setCellValue(str[i]);
                    }
                }

                FileOutputStream fileOutputStream = new FileOutputStream(xlsxFileAddress);
                workBook.write(fileOutputStream);
                fileOutputStream.close();Instant end = Instant.now();
                System.out.print("\t\ttook " + timediff(start, end) + " seconds");
                System.out.println("");

                myFile = new File(xlsxFileAddress);
            }
            else
                System.out.println("Excel already exists");








        } catch (Exception ex) {
            System.out.println(ex.getMessage() + "Exception in try");
        }
        return myFile;
    }

    public static int getCellStyle(int RowNum, int ColNum, String SheetName) throws Exception {
        ExcelWSheet = ExcelWBook.getSheet(SheetName);
        String CellData;
        try {
            Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
            if (Cell.getCellType() == CellType.STRING) {
                Cell.setCellType(CellType.STRING);
            }

            CellData = Cell.getStringCellValue();
            //            if(HSSFDateUtil.isCellDateFormatted(Cell)) {
            //                Date celldata = Cell.getDateCellValue();
            //                DateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd hh:mm:ss");
            //                CellData = dateFormat.format(celldata);
            //            }

            //            return CellData;
            return 1;
        } catch (Exception e) {
            return 999;
        }
    }


    @SuppressWarnings("static-access")
    // This method is use to write value in the excel sheet
    // This method accepts four arguments (Result, Row Number, Column Number & Sheet
    // Name)
    public static void setCellData(String Result, int RowNum, int ColNum, String SheetName, String file) throws Exception {
        try {
            fis = new FileInputStream(new File(file));
            wb = new XSSFWorkbook(fis);
            sheet = wb.getSheet(SheetName);
            row = sheet.getRow(RowNum);
            Cell = row.getCell(ColNum); //UPDATE THIS CODE - THIS METHOD WON'T WORK UNLESS FIXED

            if (Cell == null) {
                //FileOutputStream fileOut = new FileOutputStream(Constants.Path_TestData);
                //ExcelWBook.write(fileOut);
                Cell = row.createCell(ColNum);
                Cell.setCellValue(Result);
                Cell.setCellStyle(sheet.getColumnStyle(ColNum));
                //fileOut.close();
                //row = sheet.getRow(1);
                //Cell = row.getCell(Constants.Col_TestStepResult);

                //Cell.setCellValue(DriverScript.bResult);
                //fis.close();

                //System.out.println("inside IF BLOCK of setCellData");
            } else {
                Cell.setCellValue(Result);
                Cell.setCellStyle(sheet.getColumnStyle(ColNum));

                //ExcelWBook = new XSSFWorkbook(new FileInputStream(Constants.Path_TestData));
                //FileOutputStream fileOut = new FileOutputStream(Constants.Path_TestData);
                //ExcelWBook.write(fileOut);
                //Cell.setCellValue(DriverScript.bResult);
                //fileOut.close();
                //System.out.println("inside IF BLOCK of setCellData");

            }
            fis.close();

            FileOutputStream output_file = new FileOutputStream(new File("pathtoexcel"));
            wb.write(output_file);
            output_file.close();
            //	        FileInputStream fis = new FileInputStream(Constants.Path_TestData);
            //	         XSSFWorkbook wb = new XSSFWorkbook(fis);
            //	        XSSFSheet sheet = wb.getSheet(Constants.Sheet_TestSteps);
            //	         Row = sheet.getRow(1);
            //	         Cell = Row.getCell(Constants.Col_TestStepResult);
            //
            //	          Cell.setCellValue(DriverScript.bResult);
            // Constant variables Test Data path and Test Data file name
            // fileOut.flush();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // This method is to get the row count used of the excel sheet
    public static int getRowCount(String SheetName) {
        ExcelWSheet = ExcelWBook.getSheet(SheetName);
        int number = ExcelWSheet.getLastRowNum() + 1;
        return number;
    }

    public static int getColumnCount(String SheetName, int rowNum) {
        int number = 0;
        for (int i = 0; i < 5; i++) {
            try {
                ExcelWSheet = ExcelWBook.getSheet(SheetName);
                number = ExcelWSheet.getRow(i).getLastCellNum();
                if (number >= 3)
                    break;
            } catch (NullPointerException e) {
                //            e.printStackTrace();
            }
        }

        return number;
    }

}
