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

import com.viiyue.plugins.excel.converter.DefaultStyleable;
import com.viiyue.plugins.excel.converter.Styleable;

@Documented
@Target( { ElementType.TYPE } )
@Retention( RetentionPolicy.RUNTIME )
public @interface Excel {

	int headerIndex() default 0;
	int startIndex() default 1;
	
	int cellHeight() default 25;
	
	boolean strip() default true;
	
	@SuppressWarnings( "rawtypes" )
	Class<? extends Styleable> styleable() default DefaultStyleable.class;
	
}
