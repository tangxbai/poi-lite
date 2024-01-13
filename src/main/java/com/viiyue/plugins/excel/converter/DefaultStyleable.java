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

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;

import com.viiyue.plugins.excel.enums.Alignment;
import com.viiyue.plugins.excel.metadata.Style;

/**
 * @author tangxbai
 * @since 1.0.0
 */
public class DefaultStyleable<T> implements Styleable<T> {

	@Override
	public Style beautifyHeader( Workbook wb, String label ) {
		return Style.of( wb, "header" ).font( "Microsoft YaHei", 10, true ).bgColor( "#D8D8D8" );
	}

	@Override
	public Style beautifyIt( Workbook wb, String label, Object value, T element, Integer num ) {
		if ( num % 2 == 0 ) {
			return Style.of( wb, "cell:row:" + ( num % 2 ) ).bgColor( "#F5F5F5" );
		}
		return Style.of( wb, "cell" );
	}

	@Override
	public void applyAll( Style style, String label ) {
		style.alignment( Alignment.LEFT_CENTER );
		if ( style.is( "header" ) ) {
			style.border( BorderStyle.THIN, "#BFBFBF" );
		} else {
			style.border( BorderStyle.HAIR, "#BFBFBF" );
		}
	}

}
