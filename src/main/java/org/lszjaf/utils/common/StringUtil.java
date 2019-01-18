package org.lszjaf.utils.common;

/**
 * @author Joybana
 * @date 2019-01-17 11:04:26
 */
public class StringUtil {
    /**
     * to upper first letter
     * @param string
     * @return
     */
    public static String toUpperFirstLetter(String string) {
        if (isEmptyOrNull(string)) {
            return string;
        }
        StringBuilder stringBuilder = new StringBuilder(string);
        stringBuilder.replace(0, 1, string.substring(0,1).toUpperCase());
        return stringBuilder.toString();
    }

    /**
     * check string is null or empty
     * isEmpty:只是判断字符串的长度是否为0
     *
     * @param string
     * @return
     */
    public static boolean isEmptyOrNull(String string) {
        if (string == null || string.trim().isEmpty()) {
            return true;
        }
        return false;
    }
}
