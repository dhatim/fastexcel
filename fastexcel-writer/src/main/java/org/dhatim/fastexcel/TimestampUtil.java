/*
 * Copyright 2016 Dhatim.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.dhatim.fastexcel;

import java.time.*;
import java.time.chrono.ChronoZonedDateTime;
import java.time.temporal.ChronoUnit;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

/**
 * Excel timestamp utility methods. For more information, check
 * <a href="https://support.office.com/en-us/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252">this</a>
 * out.
 */
public final class TimestampUtil {

    private static final int BAD_DATE = -1;
    public static final int SECONDS_PER_MINUTE = 60;
    public static final int MINUTES_PER_HOUR = 60;
    public static final int HOURS_PER_DAY = 24;
    public static final int SECONDS_PER_DAY = (HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE);
    public static final long DAY_MILLISECONDS = SECONDS_PER_DAY * 1000L;

    /**
     * Convert a {@link Date} to a serial number. Note Excel timestamps do not
     * carry any timezone information; this method uses the system timezone to
     * convert the timestamp to a serial number. If you need a specific
     * timezone, prefer using
     * {@link #convertZonedDateTime(java.time.ZonedDateTime)}.
     *
     * @param date Date value.
     * @return Serial number value.
     */
    public static Double convertDate(Date date) {
        Calendar calStart = Calendar.getInstance();
        calStart.setTime(date);
        int year = calStart.get(Calendar.YEAR);
        int dayOfYear = calStart.get(Calendar.DAY_OF_YEAR);
        int hour = calStart.get(Calendar.HOUR_OF_DAY);
        int minute = calStart.get(Calendar.MINUTE);
        int second = calStart.get(Calendar.SECOND);
        int milliSecond = calStart.get(Calendar.MILLISECOND);
        return internalGetExcelDate(year, dayOfYear, hour, minute, second, milliSecond);
    }

    public static Double convertDate(LocalDateTime date) {
        int year = date.getYear();
        int dayOfYear = date.getDayOfYear();
        int hour = date.getHour();
        int minute = date.getMinute();
        int second = date.getSecond();
        int milliSecond = date.getNano() / 1_000_000;
        return internalGetExcelDate(year, dayOfYear, hour, minute, second, milliSecond);
    }

    /**
     * Convert a {@link LocalDate} to a serial number.
     *
     * @param date Local date value.
     * @return Serial number value.
     */
    public static Double convertDate(LocalDate date) {
        int year = date.getYear();
        int dayOfYear = date.getDayOfYear();
        int hour = 0;
        int minute = 0;
        int second = 0;
        int milliSecond = 0;
        return internalGetExcelDate(year, dayOfYear, hour, minute, second, milliSecond);
    }

    /**
     * Convert a {@link ChronoZonedDateTime} to a serial number.
     *
     * @param zdt Date and timezone values.
     * @return Serial number value.
     */
    public static Double convertZonedDateTime(ZonedDateTime zdt) {
        return convertDate(zdt.toLocalDateTime());
    }

    private static double internalGetExcelDate(int year, int dayOfYear, int hour, int minute, int second, int milliSecond) {
        if (year < 1900) {
            return BAD_DATE;
        }

        // Because of daylight time saving we cannot use
        //     date.getTime() - calStart.getTimeInMillis()
        // as the difference in milliseconds between 00:00 and 04:00
        // can be 3, 4 or 5 hours but Excel expects it to always
        // be 4 hours.
        // E.g. 2004-03-28 04:00 CEST - 2004-03-28 00:00 CET is 3 hours
        // and 2004-10-31 04:00 CET - 2004-10-31 00:00 CEST is 5 hours
        double fraction = (((hour * 60.0 + minute) * 60.0 + second) * 1000.0 + milliSecond) / DAY_MILLISECONDS;

        double value = fraction + absoluteDay(year, dayOfYear);

        if (value >= 60) {
            value++;
        }

        return value;
    }

    private static int absoluteDay(int year, int dayOfYear) {
        return dayOfYear + daysInPriorYears(year);
    }

    static int daysInPriorYears(int yr) {
        if (yr < 1900) {
            throw new IllegalArgumentException("'year' must be 1900 or greater");
        }
        int yr1 = yr - 1;
        int leapDays = yr1 / 4   // plus julian leap days in prior years
                - yr1 / 100 // minus prior century years
                + yr1 / 400 // plus years divisible by 400
                - 460;      // leap days in previous 1900 years

        return 365 * (yr - 1900) + leapDays;
    }

}
