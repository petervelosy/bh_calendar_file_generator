package com.braininghub.calendarfilegenerator;

import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class BHCalendarFileGenerator {

    private static final Logger LOGGER = Logger.getLogger(BHCalendarFileGenerator.class.getName());

    public static void main(String[] args) {

        if (args.length < 3) {
            LOGGER.warning("3 arguments are expected: sourceFilePath, destFilePath, teacherName.");
        } else {

            String sourceFilePath = args[0];
            String destFilePath = args[1];
            String teacherName = args[2];

            BHExcelToICalendarConverter conv = new BHExcelToICalendarConverter();
            try {
                conv.convertBHExcelToICalendar(sourceFilePath, destFilePath, teacherName);
            } catch (IOException ex) {
                LOGGER.log(Level.SEVERE, null, ex);
            }
        }
    }

}
