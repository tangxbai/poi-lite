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
package com.viiyue.plugins.excel.enums;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * @author tangxbai
 * @since 1.0.0
 */
public enum Alignment {

	LEFT_TOP( HorizontalAlignment.LEFT, VerticalAlignment.TOP ),
	LEFT_CENTER( HorizontalAlignment.LEFT, VerticalAlignment.CENTER ),
	LEFT_BOTTOM( HorizontalAlignment.LEFT, VerticalAlignment.BOTTOM ),
	TOP( HorizontalAlignment.CENTER, VerticalAlignment.TOP ),
	RIGHT_TOP( HorizontalAlignment.RIGHT, VerticalAlignment.TOP ),
	RIGHT_CENTER( HorizontalAlignment.RIGHT, VerticalAlignment.CENTER ),
	RIGHT_BOTTOM( HorizontalAlignment.RIGHT, VerticalAlignment.BOTTOM ),
	BOTTOM( HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM ),
	CENTER( HorizontalAlignment.CENTER, VerticalAlignment.CENTER );

	HorizontalAlignment h;
	VerticalAlignment v;

	private Alignment( HorizontalAlignment h, VerticalAlignment v ) {
		this.h = h;
		this.v = v;
	}

	public void setAlignment( CellStyle style ) {
		if ( style != null ) {
			style.setAlignment( h );
			style.setVerticalAlignment( v );
		}
	}

}
