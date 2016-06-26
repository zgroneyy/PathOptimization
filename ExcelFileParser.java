package com.company;
/*
A java program to get rid of time-consuming copy-paste operations among MS Excel files.
Author: Özgür Öney
Republic of Turkey Ministry of Forest & Water Management, Jun 2016.

IMPORTANT NOTE:
.xls -> File extension for versions of MS EXCEL released before 2007.
.xlsx -> File Extension for versions of MS EXCEL released after 2007.

Main difference between manipulating .xls and .xlsx is external libraries used. While .xlsx uses Apache POI XSSF, .xls uses APACHE POI HSSF.
 */

import java.io.*;
import java.lang.*;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.*;

public class ExcelFileParser {
    private static Logger logger = Logger.getLogger(ExcelFileParser.class.getPackage().getName());
    // For xlsx type
    private XSSFWorkbook workBookXlsx;
    private XSSFSheet workBookSheetXlsx;
    private static String[][] bigpic = new String[375][3];
    // For xls type
    private HSSFWorkbook workBookXls;
    private HSSFSheet workBookSheetXls;
    //Constructor for default named sheets
    public ExcelFileParser(String readpath, String writepath, boolean isXlsx) throws Exception {
        for(int i=0; i<375; i++ )
            for(int j=0; j<3; j++)
                bigpic[i][j]="X";
        if(isXlsx) {
            String[] files= listFilesForFolder(readpath);
            for(int i=0; i<files.length;i++){
                //System.out.println("Alınan dosyalar: " + files[i]);
                //System.out.println("Tüm dosyalar başarıyla alındı!");
            }
            for(int i=0; i<files.length;i++){
                bigpic[i]=parseExcelFileIntoArr(readpath + "/" + files[i], true);
            }
            //for(int i=0; i<bigpic.length; i++)
              //  System.out.println("Bigpic array" + bigpic[i][1]);
            writeIntoExcelFile(bigpic, writepath);
            //If documents is MS Excel 2007+, means .xlsx format. Behave accordingly.
        }else {
            String[] files= listFilesForFolder(readpath);
            for(int i=0; i<files.length;i++){
                bigpic[i]=parseExcelFileIntoArr(readpath + files[i], false);
            }
            writeIntoExcelFile(bigpic, writepath);
        }
    }
    //Constructor for customized-named sheets in excel documents.
    /*
    public ExcelFileParser(String fileName, String sheetName, boolean isXlsx) throws Exception {
        if(isXlsx) {
            workBookXlsx = new XSSFWorkbook(new FileInputStream(fileName));
            workBookSheetXlsx = workBookXlsx.getSheet(sheetName);
        }else {
            workBookXls = new HSSFWorkbook(new FileInputStream(fileName));
            workBookSheetXls = workBookXls.getSheet(sheetName);
        }
    }*/

    //Parse "Günlük Çalışma Formu" into an array, therefore manipulations on data taken from array is way easier.
    /*
    NOTE: Since time complexity is NOT so important, this way is preferred. Main aim in project is keeping data safe & secure,
    therefore time complexity is on the back burner.
    */
    public String[] listFilesForFolder(String location) {
        final File folder = new File(location);
        int counter=0;
        String extensioncheck= "";
        for (final File fileEntry : folder.listFiles()) {
            if(fileEntry.getName().length()>4)
                extensioncheck=fileEntry.getName().substring(fileEntry.getName().length()-5, fileEntry.getName().length());
            else
                extensioncheck="";
            if(extensioncheck.equals(".xlsx") && !fileEntry.getName().equals("writetest.xlsx") )
            {
                if(fileEntry.getName().charAt(0)=='~')//Indicates an open file, we dont want to read both file's itself and it's open form.
                    ;
                else
                    counter++;
            }
        }
        String[] gecici = new String[counter];
        counter=0;
        for (final File fileEntry : folder.listFiles()) {
            if(fileEntry.getName().length()>4)
                extensioncheck=fileEntry.getName().substring(fileEntry.getName().length()-5, fileEntry.getName().length());
            else
                extensioncheck="";
            if(extensioncheck.equals(".xlsx") && !fileEntry.getName().equals("writetest.xlsx"))
            {
                gecici[counter]=fileEntry.getName();
                counter++;
            }
        }
        return gecici;
    }
    private String timeConverter(double currentTime ){
        currentTime=currentTime*24;
        int time=(int)currentTime;
        currentTime=currentTime-time;
        double temporary= currentTime*60;
        temporary=Math.round(temporary);
        int minutes=(int) temporary;
        String something="";
        if(minutes==60){
            if(time<10) {
                time++;
                return "0" + time + ":" + "00"  + ":00";
            }
            else{
                time++;
                return time + ":" + "00" + ":00";
            }
        }
        else {
            if(minutes<10)
                something="0" + minutes;
            else
                something= minutes + "";
            if (time < 10)
                return "0" + time + ":" + something+ ":00";
            else
                return time + ":" + something + ":00";
        }
    }

    public String[] parseExcelFileIntoArr(String path, boolean isXlsx) throws Exception{
        workBookXlsx = new XSSFWorkbook(new FileInputStream(path));
        workBookSheetXlsx = workBookXlsx.getSheetAt(0);
        //Create a trivial array to store all information taken from .xlsx doc
        String[][] arr = new String[workBookSheetXlsx.getPhysicalNumberOfRows()][workBookSheetXlsx.getRow(0).getPhysicalNumberOfCells()];
        //Create an array to keep data to return
        String[] returnalread;
        //Define formula evaluator, which will be used in case we see a MS Excel formulated cell.
        BaseXSSFFormulaEvaluator evaluator = workBookXlsx.getCreationHelper().createFormulaEvaluator();
        //Check if the document is .xlsx
        if(isXlsx)
        {
            //Loop will scan line by line
            for (int i=0,numberOfRows = workBookSheetXlsx.getPhysicalNumberOfRows(); i < numberOfRows + 1; i++ ){
                XSSFRow row = workBookSheetXlsx.getRow(i);
                //If row is null (user did not even click on the row or did not define a known format for row)
                if(row!=null)
                {
                    //Start going right and look for respective cells.
                    for (int j = 0, numberOfColumns = row.getLastCellNum(); j < numberOfColumns; j++)
                    {
                        XSSFCell cell = row.getCell(j);
                        //Though row is not NULL, cells still can. Check whether or not.
                        if(cell!=null)
                        {
                            //To guarantee to read cell, exception is defined.
                            try {
                                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                    arr[i][j]= cell.getStringCellValue();
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_BLANK) {
                                    arr[i][j]= "";
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {
                                    arr[i][j]= String.valueOf(cell.getBooleanCellValue());
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                                    arr[i][j]= String.valueOf(cell.getNumericCellValue());
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_FORMULA) {
                                    //Initialize variables
                                    int totaltimespent=0;
                                    //Evaluate MS Excel cell's formula
                                    evaluator.evaluateFormulaCell(cell);
                                    //Now we have time in manner of DAY, so convert it into hour:minute:second manner.
                                    arr[i][j]= timeConverter(cell.getNumericCellValue());
                                } else {
                                    arr[i][j]= cell.getStringCellValue();
                                }
                            }//Thing to do whenever we cannot read a cell due to format problems.
                            catch (Exception e) {
                                logger.fatal("Oops! Can't read cell (row = " + i + ", column = " + j + ") in the excel file! Change cell format to 'Text', please!");
                            }
                        }//If row is NULL; fill it up with "" to guarantee it's string form.
                        else {
                            try{arr[i][j]= "";}
                            catch (ArrayIndexOutOfBoundsException a )
                            {
                                System.out.println("Sorunlarla karşılaşıldı. Lütfen belge içeriğini kontrol ediniz.");
                                break;
                            }


                        }
                    }
                }
            }
        }
        else
        {
            //For MS Excel version's before than 2007, .xls
            workBookXls = new HSSFWorkbook(new FileInputStream(path));
            workBookSheetXls = workBookXls.getSheetAt(0);
            return null;
        }
        //We got all data in our trivial array. Now we have to select useful data among any other in array.
        returnalread= new String[3];
        String temporary= String.valueOf(arr[0][11]);

        //Divide "Çalışma tarihi: xx.xx.xx" format into just date
        String[] dummy= temporary.split(":");
        returnalread[0]=dummy[1];
        System.out.println(returnalread[0]);
        returnalread[1]=arr[5][5];
        System.out.println(returnalread[1]);
        returnalread[2]=arr[10][5];
        System.out.println(returnalread[2]);
        return returnalread;
    }
    /*
    Method that works in same way /w method above, but instead of keeping data in array, this one returns list of lists of strings.
    Since array manipulation (walking around array and having limited memory) could be trouble for some situations, this method was written to be used
    in such scenarios by considering special request of supervisor.
     */
    public List<List<String>> parseExcelFile(boolean isXlsx) throws Exception {
        //List of lists, makes it possible to have a structure like ragged arrays.
        List<List<String>> parsedExcelFile = new ArrayList<List<String>>();
        //same procedure above, if file is .xlsx, branch here.
        if(isXlsx) {
            for (int i = 0, numberOfRows = workBookSheetXlsx.getPhysicalNumberOfRows(); i < numberOfRows + 1; i++) {
                XSSFRow row = workBookSheetXlsx.getRow(i);
                if (row != null) {
                    List<String> parsedExcelRow = new ArrayList<String>();
                    for (int j = 0, numberOfColumns = row.getLastCellNum(); j < numberOfColumns; j++) {
                        XSSFCell cell = row.getCell(j);
                        if (cell != null) {
                            try {
                                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                                    parsedExcelRow.add(cell.getStringCellValue());
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_BLANK) {
                                    parsedExcelRow.add("");
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {
                                    parsedExcelRow.add(String.valueOf(cell.getBooleanCellValue()));
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                                    parsedExcelRow.add(String.valueOf(cell.getNumericCellValue()));
                                } else if (cell.getCellType() == XSSFCell.CELL_TYPE_FORMULA) {
                                    parsedExcelRow.add("");
                                } else {
                                    parsedExcelRow.add(cell.getStringCellValue());
                                }
                            } catch (Exception e) {
                                logger.fatal("Oops! Can't read cell (row = " + i + ", column = " + j + ") in the excel file! Change cell format to 'Text', please!");
                                return null;
                            }
                        } else {
                            parsedExcelRow.add("");
                        }
                    }
                    parsedExcelFile.add(parsedExcelRow);
                }
            }
            //If file is an .xls document, we will branch here.
        }else {
            for (int i = 0, numberOfRows = workBookSheetXls.getPhysicalNumberOfRows(); i < numberOfRows + 1; i++) {
                HSSFRow row = workBookSheetXls.getRow(i);
                if (row != null) {
                    List<String> parsedExcelRow = new ArrayList<String>();
                    for (int j = 0, numberOfColumns = row.getLastCellNum(); j < numberOfColumns; j++) {
                        HSSFCell cell = row.getCell(j);
                        if (cell != null) {
                            try {
                                if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                                    parsedExcelRow.add(cell.getStringCellValue());
                                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                                    parsedExcelRow.add("");
                                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
                                    parsedExcelRow.add(String.valueOf(cell.getBooleanCellValue()));
                                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                                    parsedExcelRow.add(String.valueOf(cell.getNumericCellValue()));
                                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
                                    parsedExcelRow.add(String.valueOf(""));
                                } else {
                                    parsedExcelRow.add(cell.getStringCellValue());
                                }
                            } catch (Exception e) {
                                logger.fatal("Oops! Can't read cell (row = " + i + ", column = " + j + ") in the excel file! Change cell format to 'Text', please!");
                                return null;
                            }
                        } else {
                            parsedExcelRow.add("");
                        }
                    }
                    parsedExcelFile.add(parsedExcelRow);
                }
            }
        }

        return parsedExcelFile;
    }
    /*
        Write pre-handed data to pre-determinated MS EXCEL file.
     */
    public void writeIntoExcelFile(String[][] arr, String path)
    {
        //Protection against FileNotFoundException
        try {
            FileInputStream inp = new FileInputStream(path);
            XSSFWorkbook wb = null;
            //Protection against invalid format, a .xls file cannot be accepted.
            try {
                wb = (XSSFWorkbook) WorkbookFactory.create(inp);
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
            int counterforcells=2;

            int arraysize=0;
            for(int i=0; i<arr.length;i++){
                if(arr[i][0].length() == 11 )
                    arraysize++;
            }
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            Cell cell;
            for(int i=0; i<arraysize/5; i++ ){
                for(int j=0; j<5; j++){

                    row = sheet.getRow(i+2);
                    cell = row.getCell(j+1);
                    cell = row.createCell(j+1);
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    System.out.println("Yazdirilmasi gereken deger..." + String.valueOf(arr[j][counterforcells]));
                    cell.setCellValue(String.valueOf(arr[j][counterforcells]));
                }
            }
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream(path);
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
