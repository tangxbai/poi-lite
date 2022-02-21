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

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.Objects;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;

import com.viiyue.plugins.excel.metadata.CellInfo;

/**
 * @author tangxbai
 * @email tangxbai@hotmail.com
 * @since 2021/05/28
 */
public class DefaultReadConverter extends AbstractConverter implements ReadConverter {

	@Override
	public Object readIt( CellInfo<?> info, Class<?> fieldType, CellType cellType, Cell cell ) {
		// Normal object value
		if ( fieldType == null ) {
			return readObjectValue( cellType, cell, null );
		}

		// String
		if ( isType( fieldType, CharSequence.class ) ) {
			return cell.getStringCellValue();
		}

		// Enum
		if ( isType( fieldType, EnumConverter.class ) ) {
			return readEnumValue( fieldType, cell );
		}

		// Date
		if ( isType( fieldType, Date.class ) ) {
			return readDateValue( cellType, cell );
		}

		// LocalDate
		if ( isType( fieldType, LocalDate.class ) ) {
			LocalDateTime ldtv = cell.getLocalDateTimeCellValue();
			return ldtv == null ? null : ldtv.toLocalDate();
		}

		// LocalDateTime
		if ( isType( fieldType, LocalDateTime.class ) ) {
			return cell.getLocalDateTimeCellValue();
		}

		// BigDecimal
		if ( isType( fieldType, BigDecimal.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? null : BigDecimal.valueOf( doubleValue );
		}

		// BigInteger
		if ( isType( fieldType, BigInteger.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? null : BigInteger.valueOf( doubleValue.longValue() );
		}

		// byte × Byte
		if ( isType( fieldType, byte.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? 0 : doubleValue.byteValue();
		}
		if ( isType( fieldType, Byte.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? null : doubleValue.byteValue();
		}

		// short × Short
		if ( Objects.equals( fieldType, short.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? 0 : doubleValue.shortValue();
		}
		if ( Objects.equals( fieldType, Short.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? null : doubleValue.shortValue();
		}

		// int × Integer
		if ( isType( fieldType, int.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? 0 : doubleValue.intValue();
		}
		if ( isType( fieldType, Integer.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? null : doubleValue.intValue();
		}

		// long × Long
		if ( isType( fieldType, long.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? 0 : doubleValue.longValue();
		}
		if ( isType( fieldType, Long.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? null : doubleValue.longValue();
		}

		// float × Float
		if ( Objects.equals( fieldType, float.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? 0f : doubleValue.floatValue();
		}
		if ( Objects.equals( fieldType, Float.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? null : doubleValue.floatValue();
		}

		// double × Double
		if ( Objects.equals( fieldType, double.class ) ) {
			Double doubleValue = readDoubleValue( cellType, cell );
			return doubleValue == null ? 0D : doubleValue;
		}
		if ( Objects.equals( fieldType, Double.class ) ) {
			return readDoubleValue( cellType, cell );
		}

		// boolean × Boolean
		if ( Objects.equals( fieldType, boolean.class ) ) {
			Boolean bool = readBooleanValue( cellType, info, cell );
			return bool == null ? Boolean.FALSE : bool;
		}
		if ( Objects.equals( fieldType, Boolean.class ) ) {
			return readBooleanValue( cellType, info, cell );
		}

		log.warn( "Unsupported data type conversion \"{}\"", fieldType.getName() );
		return null;
	}

	public Object readEnumValue( Class<?> fieldType, Cell cell ) {
		String cellValue = cell.getStringCellValue();
		EnumConverter<?> [] converters = ( EnumConverter<?> [] ) fieldType.getEnumConstants();
		for ( EnumConverter<?> converter : converters ) {
			if ( Objects.equals( converter.getValue(), cellValue ) ) {
				return converter;
			}
		}
		return converters[ 0 ].getDefault();
	}

	public Date readDateValue( CellType cellType, Cell cell ) {
		if ( isNumber( cellType ) ) {
			try {
				return cell.getDateCellValue();
			} catch ( Exception e ) {
				return new Date( ( long ) cell.getNumericCellValue() );
			}
		}
		if ( isString( cellType ) ) {
			return defaultString2Date( cell.getStringCellValue() );
		}
		return null;
	}

	public Boolean readBooleanValue( CellType cellType, CellInfo<?> info, Cell cell ) {
		if ( isBoolean( cellType ) ) {
			return cell.getBooleanCellValue();
		}
		if ( isNumber( cellType ) ) {
			return cell.getNumericCellValue() != 0D;
		}
		if ( isString( cellType ) ) {
			String boolanValue = cell.getStringCellValue();
			if ( info != null && info.hasCustomBoolean() ) {
				return info.getBoolean( boolanValue );
			}
			if ( StringUtils.equalsAny( boolanValue, BOOLEAN_TRUE ) ) {
				return Boolean.TRUE;
			}
			if ( StringUtils.equalsAny( boolanValue, BOOLEAN_FALSE ) ) {
				return Boolean.FALSE;
			}
		}
		return null;
	}

	public Double readDoubleValue( CellType cellType, Cell cell ) {
		if ( isNumber( cellType ) ) {
			return cell.getNumericCellValue();
		}
		if ( isString( cellType ) ) {
			return Double.parseDouble( cell.getStringCellValue() );
		}
		if ( isBoolean( cellType ) ) {
			return cell.getBooleanCellValue() ? 1D : 0D;
		}
		return null;
	}

	public Object readObjectValue( CellType cellType, Cell cell, Object defaulValue ) {
		if ( isString( cellType ) ) {
			return cell.getStringCellValue();
		}
		if ( isNumber( cellType ) ) {
			return cell.getNumericCellValue();
		}
		if ( isBoolean( cellType ) ) {
			return cell.getBooleanCellValue();
		}
		if ( isFormula( cellType ) ) {
			return cell.getCellFormula();
		}
		RichTextString richText = cell.getRichStringCellValue();
		if ( richText != null ) {
			richText.clearFormatting();
			return richText.getString();
		}
		return defaulValue;
	}

}
