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
package com.viiyue.plugins.excel.converter;

import java.text.SimpleDateFormat;
import java.util.Objects;

import org.apache.poi.ss.usermodel.CellType;

public class AbstractConverter {

	protected static final String [] BOOLEANS = { "Y", "N", "TRUE", "FALSE", "T", "F", "1", "0" };
	protected static final ThreadLocal<SimpleDateFormat> formatter = new ThreadLocal<SimpleDateFormat>() {

		@Override
		protected SimpleDateFormat initialValue() {
			return new SimpleDateFormat( "yyyy-MM-dd HH:mm:ss" );
		}

	};

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

}
