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

import org.apache.poi.ss.usermodel.Workbook;

import com.viiyue.plugins.excel.metadata.Style;

/**
 * Excel cell style beautification interface, by implementing the interface to
 * beautify the cell style.
 * 
 * @author tangxbai
 * @email tangxbai@hotmail.com
 * @since 2021/05/28
 */
public interface Styleable<T> {

	/**
	 * Implement this method to apply the style to all cells.
	 * 
	 * @param style style chain
	 * @param label excel header cell text
	 */
	default void applyAll( Style style, String label ) {
		// Ignore implementation
	}

	/**
	 * 
	 * 
	 * @param wb Excel workbook refrence
	 * @param label Cell text
	 * @return Cell style - {@link Style}
	 */
	default Style beautifyHeader( Workbook wb, String label ) {
		return null;
	}

	/**
	 * @param wb
	 * @param label
	 * @param value
	 * @param element
	 * @param num Excel row number
	 * @return
	 */
	Style beautifyIt( Workbook wb, String label, Object value, T element, Integer num );

}
