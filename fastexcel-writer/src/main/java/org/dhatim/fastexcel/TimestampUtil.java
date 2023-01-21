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

import java.time.Duration;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.chrono.ChronoZonedDateTime;
import java.time.temporal.ChronoUnit;
import java.util.Date;

/**
 * Excel timestamp utility methods. For more information, check
 * <a href="https://support.office.com/en-us/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252">this</a>
 * out.
 */
public final class TimestampUtil {

    /**
     * Epoch for timestamps.
     */
    private static final LocalDateTime EPOCH_1900 = LocalDateTime.of(1900, 1, 1, 0, 0, 0, 0);

    private TimestampUtil() {
    }

    /**
     * Convert a {@link Date} to a serial number. Note Excel timestamps do not
     * carry any timezone information; this method uses the system timezone to
     * convert the timestamp to a serial number. If you need a specific
     * timezone, prefer using
     * {@link #convertZonedDateTime(java.time.chrono.ChronoZonedDateTime)}.
     *
     * @param date Date value.
     * @return Serial number value.
     */
    public static Double convertDate(Date date) {
        return convertDate(date, ZoneId.systemDefault());
    }

    /**
     * Convert a {@link LocalDate} to a serial number.
     *
     * @param date Local date value.
     * @return Serial number value.
     */
    public static Double convertDate(LocalDate date) {
    	double e = date.getYear() == 1900 ? 1.0 : 2.0;
        return ChronoUnit.DAYS.between(EPOCH_1900.toLocalDate(), date) + e;
    }

    /**
     * Convert a {@link ChronoZonedDateTime} to a serial number.
     *
     * @param zdt Date and timezone values.
     * @return Serial number value.
     */
    public static Double convertZonedDateTime(ChronoZonedDateTime<?> zdt) {
        return convertDate(Date.from(zdt.toInstant()), zdt.getZone());
    }

    /**
     * Internal conversion method.
     *
     * @param date Date to convert.
     * @param timezone Timezone used when performing conversion.
     * @return Excel serial number.
     */
    private static Double convertDate(Date date, ZoneId timezone) {
        LocalDateTime ldt = LocalDateTime.ofInstant(Instant.ofEpochMilli(date.getTime()), timezone);
        Duration duration = Duration.between(EPOCH_1900, ldt);
        int e = ldt.getYear() == 1900 ? 1 : 2;		// 2 to compensate for Excel bugs
        return duration.getSeconds() / 86400.0 + duration.getNano() / 86400000000000.0 + e;   
    }

}
