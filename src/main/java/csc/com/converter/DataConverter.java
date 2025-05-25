package csc.com.converter;

import java.lang.reflect.Field;

/**
 *
 * @param <T>
 * @author 刺猬
 */
public interface DataConverter<T> {
    T convert(Object value, Field field) throws Exception;
}
