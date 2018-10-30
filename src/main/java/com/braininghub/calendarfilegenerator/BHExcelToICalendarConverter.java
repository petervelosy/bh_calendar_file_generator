package com.braininghub.calendarfilegenerator;

import biweekly.Biweekly;
import biweekly.ICalendar;
import biweekly.component.VEvent;
import biweekly.util.Duration;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/* 

    Row format in the BH Timetable spreadsheet: 
        [date] [..|..|..] | (([groupName: event] | [courseLeaderName] | [mentorName] | [durationHours] |))
        The block enclosed between (( )) can be repeated an arbitrary number of times.

 */
public class BHExcelToICalendarConverter {

    private static final Logger LOGGER = Logger.getLogger(BHExcelToICalendarConverter.class.getName());

    public void convertBHExcelToICalendar(String sourceFilePath, String destFilePath, String teacherName) throws FileNotFoundException, IOException {

        LOGGER.log(Level.INFO, "Processing BH Calendar from Excel file {0}. Collecting entries for teacher {1}", new Object[]{sourceFilePath, teacherName});

        File srcFile = new File(sourceFilePath);
        File destFile = new File(destFilePath);

        ICalendar ical = new ICalendar();

        try (FileInputStream fis = new FileInputStream(srcFile);) {

            XSSFWorkbook wb = new XSSFWorkbook(fis);

            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;

            int rowCount = sheet.getPhysicalNumberOfRows();

            int colCount = getMaxColCount(rowCount, sheet);

            for (int r = 0; r < rowCount; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c < colCount; c++) {
                        cell = row.getCell(c);
                        if (cell != null) {

                            if (cellContainsTeacherName(cell, teacherName)) {
                                extractCourseEventAndAddToCalendar(row, r, c, ical);
                            }
                        }

                    }
                }
            }

            Biweekly.write(ical).go(destFile);
            LOGGER.log(Level.INFO, "ICalendar file generated at {0}", destFilePath);
        }
    }

    private void extractCourseEventAndAddToCalendar(XSSFRow row, int rowIndex, int columnIndex, ICalendar ical) {

        XSSFCell dateCell = row.getCell(0);

        if (DateUtil.isCellDateFormatted(dateCell)) {

            VEvent event = new VEvent();
            String eventDescription = "BH: " + getGroupName(row, rowIndex, columnIndex);
            event.setSummary(eventDescription);
            event.setDescription(eventDescription);

            Date date = dateCell.getDateCellValue();
            Calendar cal = Calendar.getInstance();
            cal.setTime(date);
            XSSFCell durationCell = getDurationCell(row, rowIndex, columnIndex);
            if (durationCell != null) {
                setStartDateAndDuration(durationCell, cal, event);
                ical.addEvent(event);
            }
        }
    }

    private int getMaxColCount(int rowCount, XSSFSheet sheet) {
        XSSFRow row;
        int colCount = 0;
        int tmp = 0;
        // This ensures that we get the data properly even if it doesn't start from first few rows
        for (int i = 0; i < 10 || i < rowCount; i++) {
            row = sheet.getRow(i);
            if (row != null) {
                tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                if (tmp > colCount) {
                    colCount = tmp;
                }
            }
        }
        return colCount;
    }

    private void setStartDateAndDuration(XSSFCell durationCell, Calendar cal, VEvent event) {
        int durationHours = (int) durationCell.getNumericCellValue();
        if (durationHours <= 4) {
            cal.set(Calendar.HOUR_OF_DAY, 17);
        } else {
            cal.set(Calendar.HOUR_OF_DAY, 9);
            // Add lunch break:
            durationHours++;
        }

        Duration duration = new Duration.Builder().hours(durationHours).build();
        event.setDuration(duration);

        event.setDateStart(cal.getTime());
    }

    private boolean cellContainsTeacherName(XSSFCell cell, String teacherName) {
        return cell.getCellType() == CellType.STRING && (cell.getStringCellValue().equals(teacherName) || cell.getStringCellValue().equals("Mindenki"));
    }

    private XSSFCell getDurationCell(XSSFRow row, int rowIndex, int teacherNameCellIndex) {
        XSSFCell durationCell = row.getCell(teacherNameCellIndex + 1);
        if (durationCell == null || durationCell.getCellType() != CellType.NUMERIC) {
            durationCell = row.getCell(teacherNameCellIndex + 2);
        }
        if (durationCell == null || durationCell.getCellType() != CellType.NUMERIC) {
            LOGGER.log(Level.WARNING, "Lesson duration not found in row {0}. This is probably a summary row.", rowIndex + 1);
            return null;
        }
        return durationCell;
    }

    private String getGroupName(XSSFRow row, int r, int c) {
        XSSFCell groupNameCell = getGroupNameCell(row, r, c);
        return groupNameCell != null ? groupNameCell.getStringCellValue() : "Egyéb képzés";
    }

    private XSSFCell getGroupNameCell(XSSFRow row, int rowIndex, int teacherNameCellIndex) {
        XSSFCell groupNameCell = row.getCell(teacherNameCellIndex - 1);
        if (groupNameCell == null || groupNameCell.getCellType() != CellType.STRING || !isValidGroupName(groupNameCell.getStringCellValue())) {
            groupNameCell = row.getCell(teacherNameCellIndex - 2);
        }
        if (groupNameCell == null || groupNameCell.getCellType() != CellType.STRING || !isValidGroupName(groupNameCell.getStringCellValue())) {
            LOGGER.log(Level.WARNING, "Group name not found in row {0}. This row probably belongs to an externally held course or a miscellaneous event.", rowIndex + 1);
            return null;
        }
        return groupNameCell;
    }

    private boolean isValidGroupName(String name) {
        return name.startsWith("BH") || name.startsWith("JSC");
    }

}
