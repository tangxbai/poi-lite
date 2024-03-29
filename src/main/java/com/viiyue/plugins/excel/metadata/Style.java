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

import java.awt.Color;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.concurrent.ThreadLocalRandom;
import java.util.function.Predicate;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.viiyue.plugins.excel.converter.Styleable;
import com.viiyue.plugins.excel.enums.Alignment;

/**
 * Cell beautification style
 *
 * @author tangxbai
 * @since 1.0.0
 */
public class Style {

	private static final Logger LOG = LoggerFactory.getLogger( Styleable.class );

	private static final int COLOR_BEGIN = 8;
	private static final int COLOR_BOUND = IndexedColors.values().length - 1; // 64
	private static final int COLOR_SIZE = COLOR_BOUND - COLOR_BEGIN; // 56 colors
	private static final ConcurrentMap<String, Font> FONTS = new ConcurrentHashMap<>( 128 );
	private static final ConcurrentMap<String, Short> COLORS = new ConcurrentHashMap<>( 128 );
	private static final ConcurrentMap<String, Style> INSTANCES = new ConcurrentHashMap<>( 128 );

	private final Workbook wb;
	private final CellStyle style;
	private final boolean isXssf;
	private final boolean isHssf;
	private final String namespace;

	private static final String getCacheKey( Workbook wb, String cacheKey ) {
		return Integer.toHexString( wb.hashCode() ) + ":" + cacheKey;
	}

	public static final void clear( Workbook wb ) {
		Objects.requireNonNull( wb, "Workbook object cannot be null" );
		String cachePrefix = getCacheKey( wb, "" );
		Predicate<String> predicate = key -> StringUtils.startsWith( key, cachePrefix );
		FONTS.keySet().removeIf( predicate );
		COLORS.keySet().removeIf( predicate );
		INSTANCES.keySet().removeIf( predicate );
	}

	public static final Style of( Workbook wb, String namespace ) {
		Objects.requireNonNull( wb, "Workbook object cannot be null" );
		String cacheKey = getCacheKey( wb, namespace );
		Style style = INSTANCES.get( cacheKey );
		if ( style == null ) {
			INSTANCES.put( cacheKey, style = new Style( wb, namespace ) );
		}
		return style;
	}

	private Style( Workbook wb, String namespace ) {
		this.wb = wb;
		this.namespace = namespace;
		this.style = wb.createCellStyle();
		this.isHssf = wb instanceof HSSFWorkbook;
		this.isXssf = wb instanceof XSSFWorkbook;
	}

	public boolean is( String namespace ) {
		return Objects.equals( this.namespace, namespace ) || StringUtils.startsWith( this.namespace, namespace );
	}

	public Style autoWrap() {
		style.setWrapText( true );
		return this;
	}

	public Style font( int fontSize ) {
		return font( fontSize, false );
	}

	public Style font( int fontSize, boolean blod ) {
		return font( "default", null, fontSize, blod );
	}

	public Style font( String fontName, int fontSize ) {
		return font( fontName, fontSize, false );
	}

	public Style font( String fontName, int fontSize, boolean blod ) {
		return font( fontName, null, fontSize, blod );
	}

	public Style font( String fontName, String fontColor, int fontSize, boolean blod ) {
		style.setFont( newFont( fontName, fontColor, fontSize, blod ) );
		return this;
	}

	public Style fontColor( String fontColor ) {
		return font( "default", fontColor, -1, false );
	}

	public Style bgColor( String bgColor ) {
		style.setFillPattern( FillPatternType.SOLID_FOREGROUND );
		if ( isHssf ) {
			short colorIndex = getColorIndex( bgColor );
			style.setFillForegroundColor( colorIndex );
		} else if ( isXssf ) {
			XSSFCellStyle xssfStyle = ( XSSFCellStyle ) style;
			xssfStyle.setFillForegroundColor( newXssfColor( bgColor ) );
		}
		return this;
	}

	public Style alignment( Alignment alignment ) {
		if ( alignment != null ) {
			alignment.setAlignment( style );
		}
		return this;
	}

	public Style border( BorderStyle borderStyle ) {
		style.setBorderBottom( borderStyle );
		style.setBorderLeft( borderStyle );
		style.setBorderTop( borderStyle );
		style.setBorderRight( borderStyle );
		return this;
	}

	public Style border( BorderStyle borderStyle, String borderColor ) {
		border( borderStyle );
		if ( isHssf ) {
			short colorIndex = getColorIndex( borderColor );
			style.setRightBorderColor( colorIndex );
			style.setTopBorderColor( colorIndex );
			style.setBottomBorderColor( colorIndex );
			style.setLeftBorderColor( colorIndex );
		} else if ( isXssf ) {
			XSSFColor xssfColor = newXssfColor( borderColor );
			XSSFCellStyle xssfStyle = ( XSSFCellStyle ) style;
			xssfStyle.setRightBorderColor( xssfColor );
			xssfStyle.setTopBorderColor( xssfColor );
			xssfStyle.setBottomBorderColor( xssfColor );
			xssfStyle.setLeftBorderColor( xssfColor );
		}
		return this;
	}

	public final void apply( Cell cell ) {
		if ( cell != null ) {
			cell.setCellStyle( style );
		}
	}

	private XSSFColor newXssfColor( String hexColor ) {
		XSSFWorkbook xwb = ( XSSFWorkbook ) wb;
		StylesTable source = xwb.getStylesSource();
		return new XSSFColor( Color.decode( hexColor ), source.getIndexedColors() );
	}

	private Font newFont( String fontName, String fontColor, int fontSize, boolean blod ) {
		if ( fontName == null ) {
			fontName = "default";
		}
		if ( fontColor == null ) {
			fontColor = "#";
		}
		String cacheKey = getCacheKey( wb, fontName + ":" + fontSize + ":" + fontColor + ":" + blod );
		Font font = FONTS.get( cacheKey );
		if ( font == null ) {
			font = wb.createFont();
			if ( !Objects.equals( fontName, "default" ) ) {
				font.setFontName( fontName );
			}
			if ( !Objects.equals( fontColor, "#" ) ) {
				font.setColor( getColorIndex( fontColor ) );
			}
			if ( fontSize != -1 ) {
				font.setFontHeightInPoints( ( short ) fontSize );
			}
			if ( blod ) {
				font.setBold( blod );
			}
			FONTS.put( cacheKey, font );
		}
		return font;
	}

	private short getColorIndex( String hexColor ) {
		if ( COLORS.size() >= COLOR_SIZE ) {
			LOG.warn( "Custom color value is out of range, only {} color values are allowed", COLOR_SIZE );
			return IndexedColors.AUTOMATIC.getIndex();
		}
		if ( StringUtils.isEmpty( hexColor ) ) {
			return IndexedColors.AUTOMATIC.getIndex();
		}

		String cacheKey = getCacheKey( wb, hexColor );
		Short colorIndex = COLORS.get( cacheKey );
		if ( colorIndex == null ) {
			int count = 0;
			short index = -1;
			do {
				if ( count ++ >= 10 ) {
					index = IndexedColors.AUTOMATIC.getIndex();
					break;
				}
				index = ( short ) ThreadLocalRandom.current().nextInt( COLOR_BEGIN, COLOR_BOUND );
			} while ( COLORS.containsValue( index ) );
			Color color = Color.decode( hexColor );
			byte r = ( byte ) color.getRed();
			byte g = ( byte ) color.getGreen();
			byte b = ( byte ) color.getBlue();
			HSSFPalette palette = ( ( HSSFWorkbook ) wb ).getCustomPalette();
			palette.setColorAtIndex( index, r, g, b );
			COLORS.put( cacheKey , colorIndex = index );
		}
		return colorIndex;
	}

}
