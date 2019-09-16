package cn.hiber.poi.validation;

import cn.hiber.poi.exception.HiberPoiParamException;
import org.apache.commons.lang3.StringUtils;

import java.util.Collection;

/**
 * @author hiber
 */
public class HiberPoiParamAssert {
    
    public static void notBlank(String s, String paramName) {
        if (StringUtils.isBlank(s)) {
            throw new HiberPoiParamException(paramName + " can't be blank");
        }
    }

    public static void notEmpty(CharSequence cs, String paramName) {
        if (cs == null || cs.length() == 0) {
            throw new HiberPoiParamException(paramName + " can't be empty");
        }
    }

    public static void notEmpty(Object[] array, String paramName) {
        if (array == null || array.length == 0) {
            throw new HiberPoiParamException(paramName + " can't be empty");
        }
    }
    
    public static void notEmpty(Collection<?> collection, String paramName) {
        if (collection == null || collection.isEmpty()) {
            throw new HiberPoiParamException(paramName + " can't be empty");
        }
    }

    public static void empty(CharSequence cs, String paramName) {
        if (cs != null && cs.length() > 0) {
            throw new HiberPoiParamException(paramName + " must be empty");
        }
    }

    public static void empty(Collection<?> collection, String paramName) {
        if (collection != null && !collection.isEmpty()) {
            throw new HiberPoiParamException(paramName + " must be empty");
        }
    }
    
    public static void notNull(Object object, String paramName) {
        if (object == null) {
            throw new HiberPoiParamException(paramName + " can't be null");
        }
    }
}
