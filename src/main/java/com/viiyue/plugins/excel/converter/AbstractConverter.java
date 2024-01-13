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
package com.viiyue.plugins.excel.converter;

import java.text.ParseException;
import java.util.Date;
import java.util.Objects;

import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.ss.usermodel.CellType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author tangxbai
 * @since 1.0.0
 */
public class AbstractConverter {

	protected final Logger log = LoggerFactory.getLogger( getClass() );
	protected static final String [] BOOLEAN_TRUE = { "Y", "YES", "ON", "TRUE", "T", "1" };
	protected static final String [] BOOLEAN_FALSE = { "N", "NO", "OFF", "FALSE", "F", "0" };
	protected static final String DEFAULT_DATE_FORMAT_TEMPLATE = "yyyy-MM-dd HH:mm:ss";

	public boolean isString( CellType cellType ) {
		return cellType == CellType.STRING;
	}

	public boolean isFormula( CellType cellType ) {
		return cellType == CellType.FORMULA;
	}

	public boolean isBoolean( CellType cellType ) {
		return cellType == CellType.BOOLEAN;
	}

	public boolean isNumber( CellType cellType ) {
		return cellType == CellType.NUMERIC;
	}

	public boolean isType( Class<?> source, Class<?> target ) {
		return Objects.equals( source, target ) || target.isAssignableFrom( source );
	}
	
	public Date defaultString2Date( String dateString ) {
		try {
			FastDateFormat formatter = FastDateFormat.getInstance( DEFAULT_DATE_FORMAT_TEMPLATE );
			return formatter.parse( dateString );
		} catch ( ParseException e ) {
			log.error( e.getMessage(), e );
		}
		return null;
	}
	
	public String defaultDate2String( Date date ) {
		FastDateFormat formatter = FastDateFormat.getInstance( DEFAULT_DATE_FORMAT_TEMPLATE );
		return formatter.format( date );
	}

}
