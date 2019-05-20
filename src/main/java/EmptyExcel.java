import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;


public class EmptyExcel {
    public static void main(String[] args) {
        HSSFWorkbook workBook = new HSSFWorkbook();
        HSSFSheet sheet = workBook.createSheet("Course schedule");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Name of the course");
        Cell nameOfTheCourse = row.createCell(1);
        nameOfTheCourse.setCellValue("Java basics");

        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("Hours planned");
        Cell cell3 = row2.createCell(1);
        cell3.setCellValue("9");

        Row row3 = sheet.createRow(2);
        Cell cell4;
        row3.createCell(0).setCellValue("Lesson Planned");
        Cell cell5;
        row3.createCell(1).setCellValue("6");

        Row row4 = sheet.createRow(4);
        Cell date;
        row4.createCell(0).setCellValue("Date");
        Cell dayOfTheWeek;
        row4.createCell(1).setCellValue("Day of the week");
        Cell hour1;
        row4.createCell(2).setCellValue("Start time");
        Cell hour2;
        row4.createCell(3).setCellValue("End time");

        String startHour = "18:30";
        String endHour = "20:00";
        String day1 ="Monday";
        String day2 = "Wednesday";

        Row lesson1 = sheet.createRow(5);
        Cell date1;
        lesson1.createCell(0).setCellValue("03.06.2019");
        Cell d1;
        lesson1.createCell(1).setCellValue(day1);
        Cell s1;
        lesson1.createCell(2).setCellValue(startHour);
        Cell e1;
        lesson1.createCell(3).setCellValue(endHour);

        Row lesson2 = sheet.createRow(6);
        Cell date2;
        lesson2.createCell(0).setCellValue("05.06.2019");
        Cell d2;
        lesson2.createCell(1).setCellValue(day2);
        Cell s2;
        lesson2.createCell(2).setCellValue(startHour);
        Cell e2;
        lesson2.createCell(3).setCellValue(endHour);

        Row lesson3 = sheet.createRow(7);
        Cell date3;
        lesson3.createCell(0).setCellValue("10.06.2019");
        Cell d3;
        lesson3.createCell(1).setCellValue(day1);
        Cell s3;
        lesson3.createCell(2).setCellValue(startHour);
        Cell e3;
        lesson3.createCell(3).setCellValue(endHour);

        Row lesson4 = sheet.createRow(8);
        Cell date4;
        lesson4.createCell(0).setCellValue("12.06.2019");
        Cell d4;
        lesson4.createCell(1).setCellValue(day2);
        Cell s4;
        lesson4.createCell(2).setCellValue(startHour);
        Cell e4;
        lesson4.createCell(3).setCellValue(endHour);

        Row lesson5 = sheet.createRow(9);
        Cell date5;
        lesson5.createCell(0).setCellValue("17.06.2019");
        Cell d5;
        lesson5.createCell(1).setCellValue(day1);
        Cell s5;
        lesson5.createCell(2).setCellValue(startHour);
        Cell e5;
        lesson5.createCell(3).setCellValue(endHour);

        Row lesson6 = sheet.createRow(10);
        Cell date6;
        lesson6.createCell(0).setCellValue("19.06.2019");
        Cell d6;
        lesson6.createCell(1).setCellValue(day2);
        Cell s6;
        lesson6.createCell(2).setCellValue(startHour);
        Cell e6;
        lesson6.createCell(3).setCellValue(endHour);


        for(int k = 10;k<12;k++){
            Row nextLesson = sheet.createRow(k);
            Scanner enterDate = new Scanner(System.in);
        System.out.println("Enter the date: ");
        String dateOfTheLesson = enterDate.nextLine();
        nextLesson.createCell(0).setCellValue(dateOfTheLesson);
            Scanner enterDayOfWeek = new Scanner(System.in);
            System.out.println("Enter day of week: ");
            String dayOfWeek = enterDayOfWeek.nextLine();
            nextLesson.createCell(1).setCellValue(dayOfWeek);
            Scanner enterStartHour = new Scanner(System.in);
            System.out.println("Enter start hour: ");
            String ownStartHour = enterStartHour.nextLine();
            nextLesson.createCell(2).setCellValue(ownStartHour);
            Scanner enterEndHour = new Scanner(System.in);
            System.out.println("Enter end hour: ");
            String ownEndHour = enterEndHour.nextLine();
            nextLesson.createCell(3).setCellValue(ownEndHour);
        }

        try {
            FileOutputStream out =
                    new FileOutputStream(new File("CourseSchedule.xls"));
            workBook.write(out);
            out.close();
            System.out.println("Excel written seccessfully..");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
