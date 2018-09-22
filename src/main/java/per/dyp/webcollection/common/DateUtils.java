package per.dyp.webcollection.common;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtils {
    public static final String yyyyMMddHHmmss = "yyyyMMddHHmmss";
    public static final String HH_mm_ss = "HH:mm:ss";

    private DateUtils() {

    }

    public static final String getCurrentDateTime(String format) {
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(date);
    }

    public static final String getCurrentTime(String format) {
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(date);
    }
}
