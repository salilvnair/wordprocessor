package com.github.salilvnair.wordprocessor.util;

import org.apache.commons.lang3.ClassUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;
import java.util.*;

public class AnnotationUtil {

    public static Set<Annotation> findAllAnnotations(final Class<?> cls) {
        List<Class<?>> allTypes = ClassUtils.getAllSuperclasses(cls);
        allTypes.addAll(ClassUtils.getAllInterfaces(cls));
        allTypes.add(cls);

        Set<Annotation> anns = new HashSet<Annotation>();
        for (Class<?> type : allTypes) {
            anns.addAll(Arrays.asList(type.getDeclaredAnnotations()));
        }

        Set<Annotation> superAnnotations = new HashSet<Annotation>();
        for (Annotation ann : anns) {
            findSuperAnnotations(ann.annotationType(), superAnnotations);
        }

        anns.addAll(superAnnotations);

        return anns;
    }

    private static <A extends Annotation> void findSuperAnnotations(Class<A> annotationType, Set<Annotation> visited) {
        Annotation[] anns = annotationType.getDeclaredAnnotations();

        for (Annotation ann : anns) {
            if (!ann.annotationType().getName().startsWith("java.lang") && visited.add(ann)) {
                findSuperAnnotations(ann.annotationType(), visited);
            }
        }
    }

    public static <T extends Annotation> Set<Field> findAnnotatedPublicFields(Class<? extends Object> clazz,
                                                                             Class<T> annotation) {

        if (Object.class.equals(clazz)) {
            return Collections.emptySet();
        }

        Set<Field> annotatedFields = new HashSet<Field>();
        Field[] fields = clazz.getFields();

        for (Field field : fields) {
            if (field.getAnnotation(annotation) != null) {
                annotatedFields.add(field);
            }
        }

        return annotatedFields;
    }

    public static <T extends Annotation> Set<Field> findAnnotatedFields(Class<? extends Object> clazz,
                                                                        Class<T> annotation) {
        if (Object.class.equals(clazz)) {
            return Collections.emptySet();
        }
        Set<Field> annotatedFields = new LinkedHashSet<>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (field.getAnnotation(annotation) != null) {
                annotatedFields.add(field);
            }
        }
        annotatedFields.addAll(findAnnotatedFields(clazz.getSuperclass(), annotation));
        return annotatedFields;
    }


    public static <T extends Annotation> Set<Method> getAnnotatedPublicMethods(Class<?> clazz, Class<T> annotation) {
        if (Object.class.equals(clazz)) {
            return Collections.emptySet();
        }

        List<Class<?>> ifcs = ClassUtils.getAllInterfaces(clazz);
        Set<Method> annotatedMethods = new HashSet<Method>();

        Method[] methods = clazz.getMethods();

        for (Method method : methods) {
            if (method.getAnnotation(annotation) != null || searchOnInterfaces(method, annotation, ifcs)) {
                annotatedMethods.add(method);
            }
        }

        return annotatedMethods;
    }

    private static <T extends Annotation> boolean searchOnInterfaces(Method method, Class<T> annotationType,
                                                                     List<Class<?>> ifcs) {
        for (Class<?> iface : ifcs) {
            try {
                Method equivalentMethod = iface.getMethod(method.getName(), method.getParameterTypes());
                if (equivalentMethod.getAnnotation(annotationType) != null) {
                    return true;
                }
            } catch (NoSuchMethodException ex) {

            }
        }
        return false;
    }


    public static boolean equals(Annotation annotation,Class<?> clazz) {
        return annotation.annotationType().equals(clazz);
    }

    /**
     * Changes the annotation value for the given key of the given annotation to newValue and returns
     * the previous value.
     */
    @SuppressWarnings("unchecked")
    public static Object changeValue(Annotation annotation, String annotationKey, Object newValue){
        Object handler = Proxy.getInvocationHandler(annotation);
        Field f;
        try {
            f = handler.getClass().getDeclaredField("memberValues");
        }
        catch (NoSuchFieldException | SecurityException e) {
            throw new IllegalStateException(e);
        }
        f.setAccessible(true);
        Map<String, Object> memberValues;
        try {
            memberValues = (Map<String, Object>) f.get(handler);
        }
        catch (IllegalArgumentException | IllegalAccessException e) {
            throw new IllegalStateException(e);
        }
        Object oldValue = memberValues.get(annotationKey);
        if (oldValue == null || oldValue.getClass() != newValue.getClass()) {
            throw new IllegalArgumentException();
        }
        memberValues.put(annotationKey,newValue);
        return oldValue;
    }

}