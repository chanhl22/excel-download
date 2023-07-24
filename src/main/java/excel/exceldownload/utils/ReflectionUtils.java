package excel.exceldownload.utils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class ReflectionUtils {

    public static List<Field> getAllFields(Class<?> clazz) {
        return Arrays.stream(clazz.getDeclaredFields()).collect(Collectors.toList());
    }

    public static Field getField(Class<?> clazz, String name) throws Exception {
        for (Field field : clazz.getDeclaredFields()) {
            if (field.getName().equals(name)) {
                return clazz.getDeclaredField(name);
            }
        }
        throw new NoSuchFieldException();
    }

    public static Annotation getAnnotation(Class<?> clazz,
                                           Class<? extends Annotation> targetAnnotation) {
        if (clazz.isAnnotationPresent(targetAnnotation)) {
            return clazz.getAnnotation(targetAnnotation);
        }
        return null;
    }
}
