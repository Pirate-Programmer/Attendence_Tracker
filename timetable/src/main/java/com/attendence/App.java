package com.attendence;

import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

 
public class App 
{
    static int attendence_row_limit = 0;
    /*
    *@param
    *record: attendence workbook
    *attendence: sheet in the record workbook
    *subject: used to find the row to be modified
    *cell: the cell to be modified
     */    
    static void setAttendence(Workbook record,XSSFSheet attendence, String subject, int cell) throws FileNotFoundException, IOException
    {
        for(int row = 0; row <= attendence_row_limit; row++)
        {
            if(attendence.getRow(row).getCell(0).toString().equals(subject))
            {
                double value = attendence.getRow(row).getCell(cell).getNumericCellValue();
                attendence.getRow(row).getCell(cell).setCellValue(++value);
                record.write(new FileOutputStream("D:\\time\\timetable\\Book1.xlsx"));
            }
        }
    }

    public static void main( String[] args ) throws Exception
    {
        //fileinputStream ie takes input from file to program
        //i am READING the workbook ie the excel file
        Workbook time_table_Workbook = new XSSFWorkbook(new FileInputStream("D:\\time\\timetable\\TIMETABLE.xlsx"));
        
        //now reading the sheet in the workbook
        //as timetable is in sheet1 of workbook
        XSSFSheet sheet = (XSSFSheet)time_table_Workbook.getSheet("Sheet1");

        //did the same setup for attendence workbook where i will write the data
        Workbook attendence_Workbook = new XSSFWorkbook(new FileInputStream("D:\\time\\timetable\\Book1.xlsx"));
        XSSFSheet attendence = (XSSFSheet)attendence_Workbook.getSheet("Sheet1");

        //intializing value, contains the end limit for attendence sheet
        App.attendence_row_limit = attendence.getLastRowNum();
        

        
        Scanner input = new Scanner(System.in);
        
        String toDay = java.time.LocalDate.now().getDayOfWeek().toString();
        System.out.println(toDay);


        //tracking the row number acc to the current day
        int rowIdx = -1,cellIdx = 1;
      //  System.out.println("Row Limit: "+sheet.getLastRowNum());  //gives the row index
        for(int row = 0, end = sheet.getLastRowNum(); row <= end; row++)
        {
            //iterating throught the days cell to find the current day
            if(sheet.getRow(row).getCell(0).toString().equals(toDay))
            {
                rowIdx = row;
                break;
            }
        }

        if(rowIdx == -1)
        {
            System.out.println("Invalid Day");
            System.exit(1);
        }

        //display all the lec for the current day
        //take input y/n/c
        //y == yes(increment attended & total lecs)
        //n == no(only increment total lec)
        //c == lec cancel(perform no operation)

        int attended_cell = 1;
        int total_cell = 2;
        int cell_limit = sheet.getRow(rowIdx).getLastCellNum(); //gives the actual cell number not the index
         while(cellIdx < cell_limit)
         {
            String subject = sheet.getRow(rowIdx).getCell(cellIdx).toString();
            System.out.print(subject+" [y/n/c]:");
            switch(input.next().charAt(0))
            {
                case 'y':{
                    App.setAttendence(attendence_Workbook, attendence, subject,attended_cell);
                    App.setAttendence(attendence_Workbook, attendence, subject,total_cell);
                    break;
                }
                case 'n':{
                    App.setAttendence(attendence_Workbook, attendence, sheet.getRow(rowIdx).getCell(cellIdx).toString(),total_cell);
                    break;
                }
            }
            cellIdx++;
         }

        input.close();
        time_table_Workbook.close();
        attendence_Workbook.close();
    }
}

