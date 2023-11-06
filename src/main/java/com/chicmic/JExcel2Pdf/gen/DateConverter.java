package com.chicmic.JExcel2Pdf.gen;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
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
            return inputDate;
        }
    }
    public static String findGreatestDate(String date1, String date2) {
//        System.out.println("\u001B[44m"  + date1 + " " + date2 + "\u001B[0m");
        if(date1.isEmpty() || date1 == null) {
            return date2;
        }else if (date2.isEmpty() || date2 == null) {
            return date1;
        }
        SimpleDateFormat[] dateFormats = {
                new SimpleDateFormat("dd-MM-yyyy", Locale.ENGLISH),
                new SimpleDateFormat("dd-MMM-yyyy", Locale.ENGLISH),
                new SimpleDateFormat("dd-MMMM-yyyy", Locale.ENGLISH)
        };

        for (SimpleDateFormat dateFormat : dateFormats) {
            try {
                Date d1 = dateFormat.parse(date1);
                Date d2 = dateFormat.parse(date2);

                if (d1.after(d2)) {
                    return dateFormat.format(d1);
                } else {
                    return dateFormat.format(d2);
                }
            } catch (ParseException e) {
                // Date parsing failed for this format, try the next one
            }
        }

        System.err.println("Date parsing error: Unable to parse the dates");
        return date1;
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
