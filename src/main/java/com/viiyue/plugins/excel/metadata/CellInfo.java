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
package com.viiyue.plugins.excel.metadata;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Date;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.FastDateFormat;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.viiyue.plugins.excel.common.Helper;
import com.viiyue.plugins.excel.converter.ReadConverter;
import com.viiyue.plugins.excel.converter.Styleable;
import com.viiyue.plugins.excel.converter.WriteConverter;

/**
 * @author tangxbai
 * @since 1.0.0
 */
public class CellInfo<T> {

	protected static final Logger log = LoggerFactory.getLogger( CellInfo.class );

	private int index;
	private int width;
	private boolean widthAutoSize;
	private boolean ignoreHeader;
	private final String label;
	private Styleable<T> styleable;
	private ReadConverter reader;
	private WriteConverter writer;

	private String [] bools;
	private String dateformat;

	private Class<?> beanType;
	private Field field;
	private Class<?> fieldType;
	private String fieldName;
	private Method getter;
	private Method setter;

	private CellInfo( String label ) {
		this.label = label;
	}

	public static final <T> CellInfo<T> newCell( String label ) {
		return new CellInfo<T>( label );
	}
	
	public static final CellInfo<Map<String,Object>> newMapCell( String label ) {
		return new CellInfo<Map<String,Object>>( label );
	}

	public CellInfo<T> index( int index ) {
		this.index = index;
		return this;
	}

	public CellInfo<T> field( Class<?> beanType, Field field ) {
		Objects.requireNonNull( field, "Excel cell field cannot be null" );
		this.field = field;
		this.fieldType = field.getType();
		this.fieldName = field.getName();
		this.beanType = beanType;
		return this;
	}

	public CellInfo<T> widthAutoSize( boolean widthAutoSize ) {
		this.widthAutoSize = widthAutoSize;
		return this;
	}

	public CellInfo<T> ignoreHeader( boolean ignoreHeader ) {
		this.ignoreHeader = ignoreHeader;
		return this;
	}

	public CellInfo<T> width( int width ) {
		this.width = width == 0 ? 0 : width * 256;
		return this;
	}

	public CellInfo<T> setter( Method setter ) {
		this.setter = setter;
		return this;
	}

	public CellInfo<T> getter( Method getter ) {
		this.getter = getter;
		return this;
	}

	public CellInfo<T> reader( ReadConverter reader ) {
		this.reader = reader;
		return this;
	}

	public CellInfo<T> writer( WriteConverter writer ) {
		this.writer = writer;
		return this;
	}

	public CellInfo<T> styleable( Styleable<T> styleable ) {
		this.styleable = styleable;
		return this;
	}

	public CellInfo<T> bools( String ... bools ) {
		this.bools = bools;
		return this;
	}
	
	public CellInfo<T> dateformat( String dateformat ) {
		this.dateformat = dateformat;
		return this;
	}

	public int getWidth() {
		return width;
	}

	public int getIndex() {
		return index;
	}

	public boolean isColumnAutoSize() {
		return widthAutoSize;
	}

	public boolean isIgnoreHeader() {
		return ignoreHeader;
	}

	public Styleable<T> getStyleable() {
		return styleable;
	}

	public ReadConverter getReader() {
		return reader;
	}

	public WriteConverter getWriter() {
		return writer;
	}

	public Class<?> getBeanType() {
		return beanType;
	}

	public Field getField() {
		return field;
	}

	public Class<?> getFieldType() {
		return fieldType;
	}

	public String getFieldName() {
		return fieldName;
	}

	public String getDateformat() {
		return this.dateformat;
	}
	
	public boolean hasDateformat() {
		return StringUtils.isNotEmpty( dateformat );
	}
	
	public Object formatDate( Date date ) {
		if ( date != null && hasDateformat() ) {
			return FastDateFormat.getInstance( dateformat ).format( date );
		}
		return date;
	}
	
	public Object formatTemporalAccessor( TemporalAccessor accessor ) {
		if ( accessor != null && hasDateformat() ) {
			return DateTimeFormatter.ofPattern( dateformat ).format( accessor );
		}
		return accessor;
	}
	
	public boolean hasCustomBoolean() {
		return ArrayUtils.isNotEmpty( bools );
	}

	public String getBoolean( int index ) {
		return hasCustomBoolean() && ( index == 0 || index == 1 ) ? bools[ index ] : null;
	}

	public Boolean getBoolean( String value ) {
		if ( StringUtils.isNotEmpty( value ) ) {
			if ( Objects.equals( bools[ 0 ], value ) ) {
				return Boolean.TRUE;
			}
			if ( Objects.equals( bools[ 1 ], value ) ) {
				return Boolean.FALSE;
			}
		}
		return null;
	}

	public String getBooleanValue( Object value ) {
		value = ObjectUtils.defaultIfNull( value, Boolean.FALSE );
		return ArrayUtils.isEmpty( bools ) ? value.toString() : bools[ ( Boolean ) value ? 0 : 1 ];
	}

	public String getLabel() {
		return StringUtils.defaultIfEmpty( label, fieldName );
	}

	public void setFieldValue( Object instance, Object value ) {
		Helper.setFieldValue( instance, setter, field, value );
	}

	public Object getFieldValue( Object instance ) {
		return Helper.getFieldValue( instance, getter, field );
	}

}
