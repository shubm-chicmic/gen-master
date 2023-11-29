package com.chicmic.Util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.Set;
// class to covert date to 1st November 2023 format

public class DateOperations {
    private static final Set<String> DATE_FORMATS = Set.of(
            "dd-MM-yyyy",
            "dd-MMM-yyyy",
            "dd-MMMM-yyyy"
    );

    public static SimpleDateFormat getDateFormat(String dateString) {
        for (String format : DATE_FORMATS) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(format);
                Date date = sdf.parse(dateString);
                if (date != null) {
                    return sdf;
                }
            } catch (ParseException e) {
                // Date parsing failed with this format, continue to the next one
            }
        }
        return null; // No valid format found
    }
    public static String getTodaysDate() {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        Date today = new Date();
        String todayDateStr = dateFormat.format(today);
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

    public static String findMaximumDate(String date1, String date2) {
        if (date1.isEmpty() || date1 == null) {
            return date2;
        } else if (date2.isEmpty() || date2 == null) {
            return date1;
        }
        SimpleDateFormat dateFormat1 = getDateFormat(date1);
        SimpleDateFormat dateFormat2 = getDateFormat(date2);

        if (dateFormat1 != null && dateFormat2 != null) {
            try {
                Date d1 = dateFormat1.parse(date1);
                Date d2 = dateFormat2.parse(date2);

                if (d1.after(d2)) {
                    return dateFormat1.format(d1);
                } else {
                    return dateFormat2.format(d2);
                }
            } catch (ParseException e) {
                e.printStackTrace();
                return date1;
            }
        } else {
            // Date format not determined for one or both dates
            return date1;
        }

    }

    public static String findMinimumDate(String date1, String date2) {
        if (date1.isEmpty() || date1 == null) {
            return date2;
        } else if (date2.isEmpty() || date2 == null) {
            return date1;
        }
        SimpleDateFormat dateFormat1 = getDateFormat(date1);
        SimpleDateFormat dateFormat2 = getDateFormat(date2);

        try {
            // Convert the dates to the common format
            Date d1 = dateFormat1.parse(date1);
            Date d2 = dateFormat2.parse(date2);

            if (d1.before(d2)) {
                return dateFormat1.format(d1);
            } else {
                return dateFormat2.format(d2);
            }
        } catch (ParseException e) {
            e.printStackTrace();
            return date1;
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
