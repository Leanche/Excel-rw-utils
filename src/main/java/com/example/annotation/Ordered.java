package com.example.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Has higher order than @Column
 * Sets fields to respect their class order
 * from top to bottom.
 *
 * Used for mapping Excel rows to Objects
 * and storing Objects in excel
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Ordered {
}
