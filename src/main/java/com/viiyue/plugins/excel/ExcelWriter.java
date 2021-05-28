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
package com.viiyue.plugins.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.viiyue.plugins.excel.common.Helper;
import com.viiyue.plugins.excel.converter.DefaultStyleable;
import com.viiyue.plugins.excel.converter.EnumConverter;
import com.viiyue.plugins.excel.converter.Styleable;
import com.viiyue.plugins.excel.converter.WriteConverter;
import com.viiyue.plugins.excel.metadata.CellInfo;
import com.viiyue.plugins.excel.metadata.ExcelInfo;
import com.viiyue.plugins.excel.metadata.Style;

/**
 * Excel <code>poi</code> writer, used to write data list into excel table.
 * 
 * @author tangxbai
 * @email tangxbai@hotmail.com
 * @since 2021/05/28
 */
public final class ExcelWriter<T> extends ExcelProvider<ExcelWriter<T>, T> {

	private static final String defaultSheetName = "Sheet";
	private static final Styleable<Object> defaultStyleable = new DefaultStyleable<Object>();

	private int sheetIndex = 1;
	private String lastSheet;
	private Map<String, List<T>> sheets;

	public static final ExcelWriter<Map<String, Object>> of( ExcelInfo<Map<String, Object>> excel ) {
		return new ExcelWriter<Map<String, Object>>( null, excel );
	}

	public static final <T> ExcelWriter<T> of( Class<T> beanType ) {
		return new ExcelWriter<T>( beanType, null );
	}

	protected ExcelWriter( Class<T> beanType, ExcelInfo<T> em ) {
		super( beanType, em );
	}

	public ExcelWriter<T> addSheet( String sheetName ) {
		return addSheet( sheetName, new ArrayList<T>() );
	}

	public ExcelWriter<T> addSheet( List<T> datas ) {
		return addSheet( defaultSheetName + ( sheetIndex ++ ), datas );
	}

	public ExcelWriter<T> addSheet( String sheetName, List<T> datas ) {
		this.lastSheet = sheetName;
		if ( this.sheets == null ) {
			this.sheets = new LinkedHashMap<String, List<T>>( 4 );
		}
		this.sheets.put( sheetName, datas );
		return this;
	}

	public ExcelWriter<T> append( T data ) {
		if ( this.lastSheet == null ) {
			addSheet( defaultSheetName );
		}
		List<T> sources = this.sheets.get( lastSheet );
		if ( sources == null ) {
			sources = new ArrayList<T>();
		}
		sources.add( data );
		return this;
	}

	public void writeTo( String filePath ) throws FileNotFoundException, IOException {
		Objects.requireNonNull( filePath, "The Excel output file path could not be null" );
		writeTo( new File( filePath ) );
	}

	public void writeTo( File target ) throws FileNotFoundException, IOException {
		Objects.requireNonNull( target, "The target Excel file could not be null" );
		if ( target.isDirectory() ) {
			String filePath = target.getAbsolutePath() + File.separator + System.currentTimeMillis() + ".xlsx";
			target = new File( filePath );
		}
		if ( !target.exists() ) {
			target.createNewFile();
		}
		String filePath = target.getAbsolutePath();
		String extension = StringUtils.substringAfterLast( filePath, "." );
		try ( OutputStream os = new FileOutputStream( target ) ) {
			writeTo( os, StringUtils.equalsIgnoreCase( extension, "xls" ) );
		}
	}

	public void writeTo( OutputStream os, boolean xssf ) throws IOException {
		Objects.requireNonNull( os, "Target excel output stream cannot be null" );
		Objects.requireNonNull( meta, "Excel metadata cannot be null, please initialize first" );
		try ( Workbook wb = WorkbookFactory.create( xssf ) ) {
			if ( sheets != null && meta.hasCells() ) {
				for ( Entry<String, List<T>> entry : sheets.entrySet() ) {
					String sheetName = entry.getKey();
					List<T> elements = entry.getValue();
					if ( CollectionUtils.isNotEmpty( elements ) ) {
						int headerIndex = meta.getHeaderIndex();
						int startIndex = meta.getStartIndex();
						int cellHeight = meta.getCellHeight();
						Sheet sheet = wb.createSheet( sheetName );
						createHeader( wb, sheet, headerIndex, cellHeight );
						createExcelRow( wb, sheet, elements, startIndex, cellHeight );
						fixedColumnWidth( sheet );
					}
				}
			}
			wb.write( os );
		}
	}

	private void createHeader( Workbook wb, Sheet sheet, int index, int cellHeight ) {
		Row row = sheet.createRow( index );
		row.setHeightInPoints( cellHeight );
		List<CellInfo<T>> cells = meta.getCells();
		for ( int cellIndex = 0, s = cells.size(); cellIndex < s; cellIndex ++ ) {
			CellInfo<T> cellInfo = cells.get( cellIndex );
			if ( !cellInfo.isIgnoreHeader() ) {
				Cell cell = row.createCell( cellIndex );
				cell.setCellValue( cellInfo.getLabel() );
				beautifyIt( wb, cell, cellInfo, null, null, null, true );
			}
		}
		sheet.createFreezePane( 0, index );
	}

	private void createExcelRow( Workbook wb, Sheet sheet, List<T> elements, int startIndex, int cellHeight ) {
		for ( T element : elements ) {
			Row row = sheet.createRow( startIndex ++ );
			row.setHeightInPoints( cellHeight );
			if ( isBeanType ) {
				createBeanRow( wb, row, element, startIndex );
			} else {
				if ( element instanceof Map ) {
					createMapRow( wb, row, element, startIndex );
				} else {
					log.warn( "Data element only support type of \"java.lang.Map\"" );
				}
			}
		}
	}

	private void createBeanRow( Workbook wb, Row row, T element, Integer rowNumber ) {
		List<CellInfo<T>> cells = meta.getCells();
		for ( int cellIndex = 0, s = cells.size(); cellIndex < s; cellIndex ++ ) {
			CellInfo<T> cellInfo = cells.get( cellIndex );
			Cell cell = row.createCell( cellIndex );
			if ( element != null ) {
				Object value = cellInfo.getFieldValue( element );
				wirteIt( wb, cell, cellInfo, cellInfo.getFieldType(), value );
				beautifyIt( wb, cell, cellInfo, value, element, rowNumber, false );
			} else {
				beautifyIt( wb, cell, cellInfo, null, null, rowNumber, false );
			}
		}
	}

	private void createMapRow( Workbook wb, Row row, Object element, Integer rowNumber ) {
		int cellIndex = 0;
		for ( Entry<String, Object> entry : ( ( Map<String, Object> ) element ).entrySet() ) {
			String label = entry.getKey();
			CellInfo<T> cellInfo = meta.getByLabel( label );
			if ( cellInfo != null ) {
				Object value = entry.getValue();
				Cell cell = row.createCell( cellIndex ++ );
				Class<?> valueType = value == null ? String.class : value.getClass();
				wirteIt( wb, cell, cellInfo, valueType, value );
				beautifyIt( wb, cell, cellInfo, value, ( T ) element, rowNumber, false );
			}
		}
	}

	private void fixedColumnWidth( Sheet sheet ) {
		List<CellInfo<T>> cells = meta.getCells();
		for ( int cellIndex = 0, s = cells.size(); cellIndex < s; cellIndex ++ ) {
			CellInfo<T> cellInfo = cells.get( cellIndex );
			if ( cellInfo.isColumnAutoSize() ) {
				sheet.autoSizeColumn( cellIndex );
			} else if ( cellInfo.getWidth() != 0 ) {
				sheet.setColumnWidth( cellIndex, cellInfo.getWidth() );
			}
		}
	}

	private void wirteIt( Workbook wb, Cell cell, CellInfo<T> info, Class<?> type, Object value ) {
		WriteConverter converter = info.getWriter();
		Object finalValue = null;
		if ( value != null ) {
			if ( converter != null ) {
				finalValue = converter.convert( value );
			}
			if ( finalValue == null ) {
				if ( Helper.isType( type, Boolean.class ) ) {
					finalValue = info.getBooleanValue( value );
				} else if ( Helper.isType( type, Enum.class ) && value instanceof EnumConverter ) {
					finalValue = ( ( EnumConverter<?> ) value ).getValue();
				} else {
					finalValue = value;
				}
			}
		}
		if ( finalValue != null ) {
			Class<? extends Object> valueType = finalValue.getClass();
			if ( Helper.isType( valueType, Boolean.class ) ) {
				cell.setCellValue( ( Boolean ) finalValue );
			} else if ( Helper.isType( valueType, Number.class ) ) {
				cell.setCellValue( ( ( Number ) finalValue ).doubleValue() );
			} else if ( Helper.isType( valueType, Date.class ) ) {
				cell.setCellValue( ( Date ) value );
			} else if ( Helper.isType( valueType, LocalDate.class ) ) {
				cell.setCellValue( ( LocalDate ) finalValue );
			} else if ( Helper.isType( valueType, LocalDateTime.class ) ) {
				cell.setCellValue( ( LocalDateTime ) finalValue );
			} else {
				cell.setCellValue( finalValue.toString() );
			}
		} else {
			cell.setBlank();
		}
	}

	private void beautifyIt( Workbook wb, Cell cell, CellInfo<T> info, Object value, T instance, Integer row,
			boolean isHeader ) {
		Styleable<T> styleable = info.getStyleable();
		if ( styleable != null ) {
			String label = info.getLabel();
			Style beautify = null;
			if ( isHeader ) {
				beautify = styleable.beautifyHeader( wb, label );
				if ( beautify == null ) {
					styleable = meta.getStyleable();
					if ( styleable == null ) {
						beautify = defaultStyleable.beautifyHeader( wb, label );
					} else {
						beautify = styleable.beautifyHeader( wb, label );
					}
				}
			} else {
				if ( instance != null ) {
					beautify = styleable.beautifyIt( wb, label, value, instance, row );
				}
				if ( beautify == null ) {
					styleable = meta.getStyleable();
					if ( styleable == null ) {
						beautify = defaultStyleable.beautifyIt( wb, label, value, instance, row );
					} else if ( instance != null ) {
						beautify = styleable.beautifyIt( wb, label, value, instance, row );
					}
				}
			}
			if ( beautify != null ) {
				if ( styleable == null ) {
					defaultStyleable.applyAll( beautify, label );
				} else {
					styleable.applyAll( beautify, label );
				}
				beautify.apply( cell );
			}
		}
	}

}
