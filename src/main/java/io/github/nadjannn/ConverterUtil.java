package io.github.nadjannn;

import org.apache.commons.collections4.CollectionUtils;

import java.text.NumberFormat;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Utility for process Strings and arrays.
 */
public class ConverterUtil {

    /**
     * Converts Object to String. Null value will be converted to empty string. Removes extra zero for numerical value after dot.
     *
     * @param value Object value.
     * @param format boolean value. Apply local settings for number representation if it is true or use default dot otherwise.
     * @return String representation.
     */
    public static String convertToString(Object value, boolean format) {
        if (value == null) return "";
        if (value instanceof Double) {
            return convertDoubleToString((Double) value, format);
        }
        return value.toString();
    }

    /**
     * Converts double numerical value to String. Removes extra zero after dot (or other sign if format is true and there is other character for digits separation).
     *
     * @param value  Double value.
     * @param format boolean value. Apply local settings for number representation if it is true or use default dot otherwise.
     * @return String representation.
     */
    private static String convertDoubleToString(Double value, boolean format) {
        if (value == null) return "";
        if (value.longValue() == value.doubleValue()) {
            return ((Long) value.longValue()).toString();
        }
        String stringValue = format ? NumberFormat.getInstance().format(value) : value.toString();
        if (stringValue.endsWith(".0") || stringValue.endsWith(",0")) {
            stringValue = stringValue.substring(0, stringValue.length() - 2);
        }
        return stringValue;
    }

    /**
     * Converts list of String to array of String and filters out null String values.
     *
     * @param list List of String objects.
     * @return arrays of Strings without null String values.
     */
    public static String[] convertToArrayWithoutNulls(List<String> list) {
        if (CollectionUtils.isEmpty(list)) {
            return new String[]{};
        }
        List<String> filteredList = list.stream().filter(v -> v != null).collect(Collectors.toList());
        String[] array = new String[filteredList.size()];
        return filteredList.toArray(array);
    }

}
