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
package com.viiyue.plugins.excel.common;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Objects;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Helper {

	private static final Logger log = LoggerFactory.getLogger( Helper.class );

	public static boolean isType( Class<?> source, Class<?> target ) {
		return Objects.equals( source, target ) || target.isAssignableFrom( source );
	}

	public static final void setFieldValue( Object instance, Method setter, Field field, Object value ) {
		if ( setter != null ) {
			try {
				if ( !setter.isAccessible() ) {
					setter.setAccessible( true );
				}
				setter.invoke( instance, value );
			} catch ( Exception e ) {
				log.error( e.getMessage(), e );
			}
		}
		if ( field != null ) {
			try {
				if ( !field.isAccessible() ) {
					field.setAccessible( true );
				}
				field.set( instance, value );
			} catch ( Exception e ) {
				log.error( e.getMessage(), e );
			}
		}
	}

	public static final Object getFieldValue( Object instance, Method getter, Field field ) {
		if ( getter != null ) {
			try {
				if ( !getter.isAccessible() ) {
					getter.setAccessible( true );
				}
				return getter.invoke( instance );
			} catch ( Exception e ) {
				log.error( e.getMessage(), e );
			}
		}
		if ( field != null ) {
			try {
				if ( !field.isAccessible() ) {
					field.setAccessible( true );
				}
				return field.get( instance );
			} catch ( Exception e ) {
				log.error( e.getMessage(), e );
			}
		}
		return null;
	}

}
