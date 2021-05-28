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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;

import com.viiyue.plugins.excel.converter.ReadConverter;
import com.viiyue.plugins.excel.converter.Styleable;
import com.viiyue.plugins.excel.converter.WriteConverter;

public class ExcelInfo<T> {

	private int headerIndex;
	private int startIndex;
	private int cellHeight;
	private boolean strip;
	private Styleable<T> styleable;
	private ReadConverter reader;
	private WriteConverter writer;
	private List<CellInfo<T>> cells;
	private Map<String, CellInfo<T>> mappings;

	protected ExcelInfo() {}

	public static final <T> ExcelInfo<T> build() {
		return new ExcelInfo<T>();
	}

	public ExcelInfo<T> headerIndex( int index ) {
		this.headerIndex = index;
		return this;
	}

	public ExcelInfo<T> strip( boolean strip ) {
		this.strip = strip;
		return this;
	}

	public ExcelInfo<T> cellHeight( int cellHeight ) {
		this.cellHeight = cellHeight;
		return this;
	}

	public ExcelInfo<T> reader( ReadConverter reader ) {
		this.reader = reader;
		return this;
	}

	public ExcelInfo<T> writer( WriteConverter writer ) {
		this.writer = writer;
		return this;
	}

	public ExcelInfo<T> styleable( Styleable<T> styleable ) {
		this.styleable = styleable;
		return this;
	}

	public ExcelInfo<T> startIndex( int index ) {
		this.startIndex = index;
		return this;
	}

	public ExcelInfo<T> addCell( CellInfo<T> cell ) {
		if ( cell != null ) {
			if ( this.cells == null ) {
				this.cells = new ArrayList<CellInfo<T>>();
				this.mappings = new HashMap<String, CellInfo<T>>();
			}
			if ( cell.getReader() == null ) {
				cell.reader( reader );
			}
			if ( cell.getWriter() == null ) {
				cell.writer( writer );
			}
			if ( cell.getStyleable() == null ) {
				cell.styleable( styleable );
			}
			this.cells.add( cell );
			this.mappings.put( cell.getLabel(), cell );
		}
		return this;
	}

	public ExcelInfo<T> cells( String ... labels ) {
		for ( String label : labels ) {
			addCell( CellInfo.newCell( label ) );
		}
		return this;
	}

	public CellInfo<T> getByLabel( String label ) {
		return mappings == null ? null : mappings.get( strip ? StringUtils.strip( label ) : label );
	}

	public List<CellInfo<T>> getCells() {
		return cells;
	}

	public boolean hasCells() {
		return CollectionUtils.isNotEmpty( cells );
	}

	public Styleable<T> getStyleable() {
		return styleable;
	}

	public int getHeaderIndex() {
		return headerIndex;
	}

	public int getStartIndex() {
		return startIndex == 0 ? headerIndex + 1 : startIndex;
	}

	public int getCellHeight() {
		return this.cellHeight;
	}

}
