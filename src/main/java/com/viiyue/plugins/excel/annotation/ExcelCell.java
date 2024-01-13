/**
 * Copyright (C) 2022 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.viiyue.plugins.excel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.viiyue.plugins.excel.converter.DefaultReadConverter;
import com.viiyue.plugins.excel.converter.ReadConverter;
import com.viiyue.plugins.excel.converter.Styleable;
import com.viiyue.plugins.excel.converter.WriteConverter;

/**
 * Excel cell metadata descriptor
 *
 * @author tangxbai
 * @since 1.0.0
 */
@Documented
@Target( { ElementType.FIELD, ElementType.METHOD } )
@Retention( RetentionPolicy.RUNTIME )
public @interface ExcelCell {

    /**
     * Sse field name by default
     * 
     * @return the header cell
     */
    String label();

    /**
     * Whether to ignore this content during write operation
     * 
     * @return ignore status
     */
    boolean ignoreHeader() default false;

    /**
     * <p>
     * Set the cell width, the unit is the number of characters.
     * 
     * <p>
     * <table border="true" width="80%">
     * <tr>
     * <td>TEXT</td>
     * <td align="center">WIDTH</td>
     * </tr>
     * <tr>
     * <td>测试文字</td>
     * <td align="center">4</td>
     * </tr>
     * <tr>
     * <td>Test words</td>
     * <td align="center">10</td>
     * </tr>
     * </table>
     * 
     * @return the excel cell width
     */
    int width() default 0;

    /**
     * Cell adaptive text width
     * 
     * @return whether the adaptive width is available?
     */
    boolean widthAutoSize() default false;

    /**
     * The time format pattern of Date, LocalDate and LocalDateTime.
     * 
     * @return date format pattern
     */
    String dateformat() default "yyyy-MM-dd HH:mm:ss";

    /**
     * @return boolean replacement value, fill in the order of <b>true/false</b>.
     */
    String [] bools() default {};

    /**
     * Cell style
     * 
     * @return the class of the {@link Styleable}
     */
    @SuppressWarnings( "rawtypes" )
    Class<? extends Styleable> styleable() default Styleable.class;

    /**
     * Custom cell write converter
     * 
     * @return the {@link WriteConverter} class
     */
    Class<? extends WriteConverter> writer() default WriteConverter.class;

    /**
     * Custom cell read converter
     * 
     * @return the {@link ReadConverter} class
     */
    Class<? extends ReadConverter> reader() default DefaultReadConverter.class;

}
