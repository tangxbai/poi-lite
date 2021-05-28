/**
 * Copyright (C) 2021 the original author or authors.
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

@Documented
@Target( { ElementType.FIELD, ElementType.METHOD } )
@Retention( RetentionPolicy.RUNTIME )
public @interface ExcelCell {

	/**
	 * @return use field name by default
	 */
	String label() default "";

	/**
	 * @return
	 */
	boolean strip() default false;
	
	boolean ignoreHeader() default false;

	/**
	 * @return excel cell width
	 */
	int width() default 0;

	/**
	 * @return boolean of automatic width adjustment
	 */
	boolean widthAutoSize() default false;

	/**
	 * @return boolean replacement value, fill in the order of
	 *         <b>true/false</b>.
	 */
	String [] bools() default {};

	@SuppressWarnings( "rawtypes" )
	Class<? extends Styleable> styleable() default Styleable.class;

	Class<? extends WriteConverter> writer() default WriteConverter.class;

	Class<? extends ReadConverter> reader() default DefaultReadConverter.class;

}
