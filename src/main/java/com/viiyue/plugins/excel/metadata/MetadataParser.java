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
package com.viiyue.plugins.excel.metadata;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.viiyue.plugins.excel.annotation.Excel;
import com.viiyue.plugins.excel.annotation.ExcelCell;
import com.viiyue.plugins.excel.converter.ReadConverter;
import com.viiyue.plugins.excel.converter.Styleable;
import com.viiyue.plugins.excel.converter.WriteConverter;

public class MetadataParser<T> {

	protected static final Logger log = LoggerFactory.getLogger( MetadataParser.class );
	private static final Map<Class<?>, ExcelInfo<?>> metas = new ConcurrentHashMap<Class<?>, ExcelInfo<?>>( 8 );
	private static final Map<Class<?>, Object> caches = new ConcurrentHashMap<Class<?>, Object>( 8 );

	protected final Class<?> beanType;
	protected final boolean isBeanType;

	public MetadataParser( Class<?> beanType ) {
		this.beanType = beanType;
		this.isBeanType = !Objects.equals( beanType, Object.class );
		this.parseTemplate( beanType );
	}

	public boolean isBeanType() {
		return this.isBeanType;
	}

	public ExcelInfo<T> getMetadata( Class<T> beanType ) {
		return ( ExcelInfo<T> ) metas.get( beanType );
	}

	public <K> K newInstance( Class<K> clazz ) {
		if ( clazz != null ) {
			try {
				return clazz.newInstance();
			} catch ( InstantiationException e ) {
				log.error( e.getMessage(), e );
			} catch ( IllegalAccessException e ) {
				log.error( e.getMessage(), e );
			}
		}
		return null;
	}

	private void parseTemplate( Class<?> beanType ) {
		if ( isBeanType && !metas.containsKey( beanType ) ) {
			ExcelInfo<T> meta = ExcelInfo.build();
			parseTypeAnnotation( meta );
			parseMemberAnnotation( meta );
			metas.put( beanType, meta );
		}
	}

	private void parseTypeAnnotation( ExcelInfo<T> meta ) {
		Excel excel = beanType.getAnnotation( Excel.class );
		if ( excel != null ) {
			meta.strip( excel.strip() );
			meta.headerIndex( excel.headerIndex() );
			meta.startIndex( excel.startIndex() );
			meta.cellHeight( excel.cellHeight() );
			meta.styleable( getSingleton( excel.styleable(), Styleable.class, null ) );
		}
	}

	private void parseMemberAnnotation( ExcelInfo<T> meta ) {
		for ( Field field : beanType.getDeclaredFields() ) {
			int modifiers = field.getModifiers();
			if ( Modifier.isFinal( modifiers ) && Modifier.isStatic( modifiers ) ) {
				continue;
			}

			Method getter = null;
			Method setter = null;
			try {
				PropertyDescriptor descriptor = new PropertyDescriptor( field.getName(), beanType );
				getter = descriptor.getReadMethod();
				setter = descriptor.getWriteMethod();
			} catch ( IntrospectionException e ) {
				log.warn( "Field \"{}\" in class \"{}\" has no getter method", field.getName(), beanType.getName() );
			}

			ExcelCell cell = field.getAnnotation( ExcelCell.class );
			if ( cell == null && getter != null ) {
				cell = getter.getAnnotation( ExcelCell.class );
			}

			if ( cell != null ) {
				CellInfo<T> info = CellInfo.newCell( cell.label() );
				info.field( beanType, field );
				info.width( cell.width() );
				info.bools( cell.bools() );
				info.ignoreHeader( cell.ignoreHeader() );
				info.getter( getter ).setter( setter );
				info.reader( getSingleton( cell.reader(), ReadConverter.class, null ) );
				info.writer( getSingleton( cell.writer(), WriteConverter.class, null ) );
				info.styleable( getSingleton( cell.styleable(), Styleable.class, meta.getStyleable() ) );
				info.widthAutoSize( cell.widthAutoSize() );
				meta.addCell( info );
			}
		}
	}

	private <K> K getSingleton( Class<?> cacheType, Class<K> targetType, K defaultValue ) {
		if ( cacheType == null ) {
			return defaultValue;
		}
		if ( !targetType.isInterface() || Objects.equals( cacheType, targetType ) ) {
			return defaultValue;
		}
		if ( cacheType.isInterface() ) {
			log.error( "{} cannot be an interface of \"{}\"", targetType.getSimpleName(), cacheType.getName() );
			return defaultValue;
		}
		Object object = null;
		synchronized ( caches ) {
			object = caches.get( cacheType );
			if ( object == null ) {
				object = newInstance( cacheType );
				if ( object != null ) {
					caches.put( cacheType, object );
				}
			}
		}
		if ( object != null ) {
			Class<? extends Object> valueType = object.getClass();
			if ( Objects.equals( targetType, valueType ) || targetType.isAssignableFrom( valueType ) ) {
				return targetType.cast( object );
			}
		}
		return defaultValue;
	}

}
