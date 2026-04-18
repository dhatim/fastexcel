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

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.util.Date;

/**
 * Excel timestamp utility methods. For more information, check
 * <a href="https://support.office.com/en-us/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252">this</a>
 * out.
 */
public final class TimestampUtil {

    private static final double BAD_DATE = -1;
    @Deprecated
    public static final int SECONDS_PER_MINUTE = 60;
    @Deprecated
    public static final int MINUTES_PER_HOUR = 60;
    @Deprecated
    public static final int HOURS_PER_DAY = 24;
    @Deprecated
    public static final int SECONDS_PER_DAY = (HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE);
    @Deprecated
    public static final long DAY_MILLISECONDS = SECONDS_PER_DAY * 1000L;

    private static final long DAYS_TO_MILLIS = 86_400_000L;
    private static final long EXCEL_EPOCH_MILLIS = LocalDate.of(1899, 12, 31)
            .atStartOfDay(ZoneOffset.UTC)
            .toInstant()
            .toEpochMilli();

    private static double epochMillisToExcel(long epochMillis) {
        double value = (epochMillis - EXCEL_EPOCH_MILLIS) / (double) DAYS_TO_MILLIS;
        // Excel leap year bug: serial >= 60 is off by one
        if (value >= 60) {
            value++;
        }
        return value;
    }

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
        return convertDate(date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
    }

    /**
     * Convert a {@link LocalDateTime} to a serial number.
     *
     * @param localDateTime Local date time value.
     * @return Serial number value.
     */
    public static Double convertDate(LocalDateTime localDateTime) {
        if (localDateTime.getYear() < 1900) {
            return BAD_DATE;
        }
        LocalTime time = localDateTime.toLocalTime();
        long epochMillis = localDateTime.toLocalDate().toEpochDay() * DAYS_TO_MILLIS
                + time.getHour() * 3_600_000L
                + time.getMinute() * 60_000L
                + time.getSecond() * 1_000L
                + time.getNano() / 1_000_000L;
        return epochMillisToExcel(epochMillis);
    }

    /**
     * Convert a {@link LocalDate} to a serial number.
     *
     * @param localDate Local date value.
     * @return Serial number value.
     */
    public static Double convertDate(LocalDate localDate) {
        if (localDate.getYear() < 1900) {
            return BAD_DATE;
        }
        return epochMillisToExcel(localDate.toEpochDay() * DAYS_TO_MILLIS);
    }

    /**
     * Convert a {@link ZonedDateTime} to a serial number.
     *
     * @param zonedDateTime Date and timezone values.
     * @return Serial number value.
     */
    public static Double convertZonedDateTime(ZonedDateTime zonedDateTime) {
        return convertDate(zonedDateTime.toLocalDateTime());
    }

    /**
     * Convert an {@link Instant} to a serial number. This method does care about
     * timestamp information, all conversions are done in UTC.
     *
     * @param instant Instant value.
     * @return Serial number value.
     */
    public static Double convertInstant(Instant instant) {
        return epochMillisToExcel(instant.toEpochMilli());
    }

}
