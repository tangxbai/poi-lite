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
package com.viiyue.plugins.excel;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.viiyue.plugins.excel.converter.DefaultReadConverter;
import com.viiyue.plugins.excel.converter.ReadConverter;
import com.viiyue.plugins.excel.metadata.CellInfo;
import com.viiyue.plugins.excel.metadata.ExcelInfo;

/**
 * Excel <code>poi</code> reader
 * 
 * @author tangxbai
 * @email tangxbai@hotmail.com
 * @since 2021/05/28
 */
public final class ExcelReader<T> extends ExcelProvider<ExcelReader<T>, T> {

	private static final int defaultSheetIndex = -1;
	private static final ReadConverter defaultReader = new DefaultReadConverter();

	private Map<Integer, T> dataList;

	public static final <T> ExcelReader<T> of( Class<T> beanType ) {
		return new ExcelReader<T>( beanType, null );
	}

	public static final ExcelReader<Map<String, Object>> of( ExcelInfo<Map<String, Object>> em ) {
		return new ExcelReader<Map<String, Object>>( null, em );
	}

	private ExcelReader( Class<T> beanType, ExcelInfo<T> em ) {
		super( beanType, em );
	}

	public ExcelReader<T> read( String filePath ) throws EncryptedDocumentException, IOException {
		return read( filePath, defaultSheetIndex );
	}
	
	public ExcelReader<T> read( String filePath, int sheetIndex ) throws EncryptedDocumentException, IOException {
		return read( new File( filePath ), sheetIndex );
	}
	
	public ExcelReader<T> read( File file ) throws EncryptedDocumentException, IOException {
		return read( file, defaultSheetIndex );
	}

	public ExcelReader<T> read( File file, int sheedIndex ) throws EncryptedDocumentException, IOException {
		Objects.requireNonNull( file, "The target Excel file could not be null" );
		if ( !file.exists() ) {
			throw new IOException( "The target Excel file could not be found : \"" + file.getAbsolutePath() + "\"" );
		}
		if ( !file.canRead() ) {
			throw new IOException( "Failed to read the target Excel file : \"" + file.getAbsolutePath() + "\"" );
		}
		try ( Workbook wb = WorkbookFactory.create( file ) ) {
			parseWorkbook( wb, sheedIndex );
		}
		return this;
	}

	public ExcelReader<T> read( InputStream is ) throws EncryptedDocumentException, IOException {
		return read( is, defaultSheetIndex );
	}
	
	public ExcelReader<T> read( InputStream is, int sheedIndex ) throws EncryptedDocumentException, IOException {
		Objects.requireNonNull( is, "The Excel input stream is null" );
		try ( Workbook wb = WorkbookFactory.create( is ) ) {
			parseWorkbook( wb, sheedIndex );
		}
		return this;
	}

	public boolean hasAny() {
		return dataList != null && dataList.size() > 0;
	}

	public ExcelReader<T> forEach( java.util.function.BiConsumer<Integer, T> consumer ) {
		if ( hasAny() ) {
			dataList.forEach( consumer );
		}
		return this;
	}

	public List<T> getReadList() {
		return dataList == null ? Collections.emptyList() : new ArrayList<T>( dataList.values() );
	}

	private void parseWorkbook( Workbook wb, int sheedIndex ) {
		if ( isBeanType ) {
			parseWorkbookWithBean( wb, sheedIndex );
		} else {
			parseWorkbookWithObject( wb, sheedIndex );
		}
	}

	private void parseWorkbookWithBean( Workbook wb, int sheedIndex ) {
		Objects.requireNonNull( meta, "Excel metadata cannot be null, please initialize first" );

		int activeIndex = sheedIndex == defaultSheetIndex ? wb.getActiveSheetIndex() : sheedIndex;
		Sheet sheet = wb.getSheetAt( activeIndex );
		initHeader( sheet );

		for ( int i = meta.getStartIndex(), s = sheet.getPhysicalNumberOfRows(); i <= s; i ++ ) {
			Row row = sheet.getRow( i );
			if ( row != null ) {
				T instance = mp.newInstance( beanType );
				if ( instance != null ) {
					for ( CellInfo<T> cellInfo : meta.getCells() ) {
						Cell cell = row.getCell( cellInfo.getIndex() );
						if ( cell != null ) {
							CellType cellType = cell.getCellType();
							Class<?> fieldType = cellInfo.getField().getType();
							ReadConverter reader = ObjectUtils.defaultIfNull( cellInfo.getReader(), defaultReader );
							Object cellValue = reader.readIt( cellInfo, fieldType, cellType, cell );
							cellInfo.setFieldValue( instance, cellValue );
						}
					}
				}
				addDataElement( instance, i + 1 );
			}
		}
	}

	private void parseWorkbookWithObject( Workbook wb, int sheedIndex ) {
		Objects.requireNonNull( meta, "Excel metadata cannot be empty, please call \"#metadata(ExcelMetadata)\" to initialize" );

		int activeIndex = sheedIndex == -1 ? sheedIndex : wb.getActiveSheetIndex();
		Sheet sheet = wb.getSheetAt( activeIndex );
		initHeader( sheet );

		for ( int i = meta.getStartIndex(), s = sheet.getPhysicalNumberOfRows(); i <= s; i ++ ) {
			Row row = sheet.getRow( i );
			if ( row != null ) {
				Map<String, Object> element = new HashMap<String, Object>();
				for ( CellInfo<T> cellInfo : meta.getCells() ) {
					Cell cell = row.getCell( cellInfo.getIndex() );
					if ( cell != null ) {
						ReadConverter reader = ObjectUtils.defaultIfNull( cellInfo.getReader(), defaultReader );
						Object cellValue = reader.readIt( cellInfo, null, cell.getCellType(), cell );
						element.put( cellInfo.getLabel(), cellValue );
					}
				}
				addDataElement( ( T ) element, i + 1 );
			}
		}
	}

	private void initHeader( Sheet sheet ) {
		int headerIndex = meta.getHeaderIndex();
		Row headerRow = sheet.getRow( headerIndex );
		for ( int i = headerIndex, cs = headerRow.getPhysicalNumberOfCells(); i < cs; i ++ ) {
			Cell cell = headerRow.getCell( i );
			CellType cellType = cell.getCellType();
			if ( cellType != CellType._NONE && cellType != CellType.BLANK ) {
				String label = cell.getStringCellValue();
				CellInfo<T> cellInfo = meta.getByLabel( label );
				if ( cellInfo != null ) {
					cellInfo.index( i );
				}
			}
		}
	}

	private void addDataElement( T element, int index ) {
		if ( dataList == null ) {
			this.dataList = new HashMap<Integer, T>();
		}
		if ( element != null ) {
			this.dataList.put( index, element );
		}
	}

}
