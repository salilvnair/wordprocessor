package com.github.salilvnair.wordprocessor.reflect.annotation;

import com.github.salilvnair.wordprocessor.constant.WordProcessorConstant;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
public @interface PlaceHolder {
	String value() default "";
	boolean checkbox() default false;
	boolean nonNull() default false;
	String replaceNullWith() default "  ";
}
