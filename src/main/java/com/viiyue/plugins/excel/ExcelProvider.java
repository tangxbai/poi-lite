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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.viiyue.plugins.excel.metadata.ExcelInfo;
import com.viiyue.plugins.excel.metadata.MetadataParser;

/**
 * Excel abstract provider
 * 
 * @author tangxbai
 * @email tangxbai@hotmail.com
 * @since 2021/06/01
 */
abstract class ExcelProvider<K extends ExcelProvider<?, ?>, T> {

	protected static final Logger log = LoggerFactory.getLogger( ExcelProvider.class );

	protected boolean isBeanType;
	protected Class<T> beanType;
	protected MetadataParser<T> mp;
	protected ExcelInfo<T> meta;

	protected ExcelProvider( Class<T> beanType, ExcelInfo<T> excelInfo ) {
		this.parseBeanType( beanType );
		if ( excelInfo != null ) {
			this.meta = excelInfo;
		}
	}

	private void parseBeanType( Class<T> beanType ) {
		if ( beanType != null ) {
			this.beanType = beanType;
			this.mp = new MetadataParser<T>( beanType );
			this.meta = mp.getMetadata( beanType );
			this.isBeanType = mp.isBeanType();
		}
	}

}
