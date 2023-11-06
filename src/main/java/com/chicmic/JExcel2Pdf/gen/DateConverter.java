package com.chicmic.JExcel2Pdf.gen;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
// class to covert date to 1st November 2023 format

public class DateConverter {
    public static String getTodaysDate() {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        Date today = new Date();
        String todayDateStr =  dateFormat.format(today);
        return convertDate(todayDateStr);
    }
    private static String convertDate(String inputDate) {
        try {
            SimpleDateFormat inputFormat = new SimpleDateFormat("dd-MM-yyyy"); // Correct format for your input date
            Date date = inputFormat.parse(inputDate);

            SimpleDateFormat dayFormat = new SimpleDateFormat("d");
            int day = Integer.parseInt(dayFormat.format(date));

            SimpleDateFormat monthYearFormat = new SimpleDateFormat("MMMM yyyy");
            String monthYear = monthYearFormat.format(date);

            String ordinalDay = addOrdinalIndicator(day);

            return ordinalDay + " " + monthYear;
        } catch (ParseException e) {
            e.printStackTrace();
            return inputDate; // Return the original date in case of an error
        }
    }


    private static String addOrdinalIndicator(int day) {
        if (day >= 11 && day <= 13) {
            return day + "th";
        }

        switch (day % 10) {
            case 1:
                return day + "st";
            case 2:
                return day + "nd";
            case 3:
                return day + "rd";
            default:
                return day + "th";
        }
    }

}
