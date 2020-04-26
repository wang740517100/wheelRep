package cn.wangkf.utils


import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * Created by stanley.wang on 2020/4/13.
 */
public class DateUtils {

    public static Integer moveTime(Integer timeStamp, Integer years, Integer months, Integer days,
                                    Integer hours, Integer minutes, Integer seconds) {
        Calendar curr = Calendar.getInstance();
        long longTimeStamp = (long)timeStamp * 1000;
        curr.setTime(new Date(longTimeStamp));


        curr.set(years != null ? curr.get(Calendar.YEAR) + years : curr.get(Calendar.YEAR),
                months != null ? curr.get(Calendar.MONTH) + months : curr.get(Calendar.MONTH),
                days != null ? curr.get(Calendar.DATE) + days : curr.get(Calendar.DATE),
                hours != null ? curr.get(Calendar.HOUR) + hours : curr.get(Calendar.HOUR),
                minutes != null ? curr.get(Calendar.MINUTE) + minutes : curr.get(Calendar.MINUTE),
                seconds != null ? curr.get(Calendar.SECOND) + seconds : curr.get(Calendar.SECOND));

        Date date = curr.getTime();
        int time = Math.toIntExact (date.getTime()/1000);
        return time;
    }

    public static Integer initTime(Integer timeStamp, Integer year, Integer month, Integer day,
                                   Integer hour, Integer minute, Integer second) {
        Calendar curr = Calendar.getInstance();
        long longTimeStamp = (long)timeStamp * 1000;
        curr.setTime(new Date(longTimeStamp));

        curr.set(year != null ? year : curr.get(Calendar.YEAR),
                month != null ? month-1 : curr.get(Calendar.MONTH),
                day != null ? day : curr.get(Calendar.DATE),
                hour != null ? hour : curr.get(Calendar.HOUR),
                minute != null ? minute : curr.get(Calendar.MINUTE),
                second != null ? second : curr.get(Calendar.SECOND));

        Date date = curr.getTime();
        int time = Math.toIntExact (date.getTime()/1000);
        return time;
    }


    /**
     * 将string转换为date
     * @param datetext
     * @param format
     * @return
     */
    public static Date string2Date(String datetext, String format) {
        try {
            SimpleDateFormat sdf;
            if (datetext == null || "".equals(datetext.trim())) {
                return new Date();
            }
            if (format != null) {
                sdf = new SimpleDateFormat(format);
            } else {
                datetext = datetext.replaceAll("/", "-");
                sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            }
            return sdf.parse(datetext);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     *将date转换为string
     * @param date
     * @param format
     * @return
     */
    public static String date2String(Date date, String format) {
        SimpleDateFormat sdf = format != null ? new SimpleDateFormat(format)
                : new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return sdf.format(date);
    }

    /**
     * 将date转换为秒级别时间戳
     * @param date
     * @return
     */
    public static Integer date2timeStamp(Date date) {
        Calendar curr = Calendar.getInstance();
        curr.setTime(date);
        return Math.toIntExact (date.getTime()/1000);
    }

    /**
     * 将秒级别时间戳转换为date
     * @param timeStamp
     * @return
     */
    public static Date timeStamp2date(Integer timeStamp) {
        long longTimeStamp = (long)timeStamp * 1000;
        return new Date(longTimeStamp);
    }

    /**
     * 获取当前时间：时
     * @return
     */
    public static int getHour()
    {
        return Calendar.getInstance().get(Calendar.HOUR_OF_DAY);
    }

    /**
     * 获取当前时间：年
     * @return
     */
    public static int getYear()
    {
        return Calendar.getInstance().get(Calendar.YEAR);
    }

    /**
     * 获取当前时间：月
     * @return
     */
    public static int getMonth()
    {
        return Calendar.getInstance().get(Calendar.MONTH) + 1;
    }

    /**
     * 获取当前时间：天
     * @return
     */
    public static int getDay()
    {
        return Calendar.getInstance().get(Calendar.DAY_OF_MONTH);
    }

    /**
     * 获取当前时间：分
     * @return
     */
    public static int getMimute()
    {
        return Calendar.getInstance().get(Calendar.MINUTE);
    }

    /**
     * 获取当前时间：秒
     * @return
     */
    public static int getSecond()
    {
        return Calendar.getInstance().get(Calendar.SECOND);
    }


}
