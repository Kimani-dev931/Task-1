package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

class Member {
    String id;
    String name;
    String mobileNumber;
    String emailAddress;
    String gender;
    List<String> errorReasons;

    public Member(String id, String name, String mobileNumber, String emailAddress, String gender) {
        this.id = id;
        this.name = name;
        this.mobileNumber = mobileNumber;
        this.emailAddress = emailAddress;
        this.gender = gender;
        this.errorReasons = new ArrayList<>();
    }


    public boolean isValid() {

        boolean isIdValid = id.matches("\\d+");
        boolean isMobileNumberValid = mobileNumber.matches("^(?:254|\\\\\\\\+254|0)?(7(?:(?:[12][0-9])|(?:0[0-8])|(?:9[0-2]))[0-9]{6})$");
        boolean isEmailAddressValid = emailAddress.matches(".+@.+");


        if (!isIdValid) {
            errorReasons.add("Invalid ID");
        }
        if (!isMobileNumberValid) {
            errorReasons.add("Invalid Mobile Number");
        }
        if (!isEmailAddressValid) {
            errorReasons.add("Invalid Email Address");
        }

        return isIdValid && isMobileNumberValid && isEmailAddressValid;
    }
}

public class ExcelProcessor {
    private static Workbook femaleWorkbook = new SXSSFWorkbook();
    private static Workbook maleWorkbook = new SXSSFWorkbook();
    private static Workbook invalidWorkbook = new SXSSFWorkbook();

    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();
        String csvFilePath = "C:\\Users\\kim\\Downloads\\member_details.csv";

        try (BufferedReader reader = new BufferedReader(new FileReader(csvFilePath))) {
            String line;
            int count = 1;
            while ((line = reader.readLine()) != null) {
                String[] data = line.split(",");
                Member member = new Member(data[0], data[1], data[2], data[3], data[4]);
                processMember(member);
                System.out.println("saved " + count);
                count++;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }


        saveWorkbookToFile(femaleWorkbook, "female_records.xlsx");
        saveWorkbookToFile(maleWorkbook, "male_records.xlsx");
        saveWorkbookToFile(invalidWorkbook, "invalid_records.xlsx");

        long endTime = System.currentTimeMillis();
        long totalTime = endTime - startTime;
        System.out.println("Total time taken: " + totalTime + " milliseconds");
    }

    private static void processMember(Member member) {
        if (member.isValid()) {
            Sheet sheet = getWorkbook(member.gender).getSheet(member.gender);
            if (sheet == null) {
                sheet = getWorkbook(member.gender).createSheet(member.gender);
            }

            Row row = sheet.createRow(sheet.getPhysicalNumberOfRows());
            row.createCell(0).setCellValue(member.id);
            row.createCell(1).setCellValue(member.name);
            row.createCell(2).setCellValue(member.mobileNumber);
            row.createCell(3).setCellValue(member.emailAddress);
        } else {
            Sheet invalidSheet = invalidWorkbook.getSheet("invalid");
            if (invalidSheet == null) {
                invalidSheet = invalidWorkbook.createSheet("invalid");
            }

            Row invalidRow = invalidSheet.createRow(invalidSheet.getPhysicalNumberOfRows());
            invalidRow.createCell(0).setCellValue(member.id);
            invalidRow.createCell(1).setCellValue(member.name);
            invalidRow.createCell(2).setCellValue(member.mobileNumber);
            invalidRow.createCell(3).setCellValue(member.emailAddress);
            invalidRow.createCell(4).setCellValue(member.gender);


            Cell errorReasonsCell = invalidRow.createCell(5);
            String errorReasons = String.join(", ", member.errorReasons);
            errorReasonsCell.setCellValue(errorReasons);
        }
    }

    private static Workbook getWorkbook(String gender) {
        switch (gender) {
            case "Female":
                return femaleWorkbook;
            case "Male":
                return maleWorkbook;
            default:
                return invalidWorkbook;
        }
    }

    private static void saveWorkbookToFile(Workbook workbook, String fileName) {
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
