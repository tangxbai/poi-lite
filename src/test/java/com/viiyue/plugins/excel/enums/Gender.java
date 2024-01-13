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

import com.viiyue.plugins.excel.converter.EnumConverter;

public enum Gender implements EnumConverter<Gender> {

	MALE( "男性" ), 
	FEMALE( "女性" ), 
	UNKNOWN( "未知" );

	private String text;

	private Gender( String text ) {
		this.text = text;
	}

	@Override
	public String getValue() {
		return text;
	}

	@Override
	public Gender getDefault() {
		return UNKNOWN;
	}

}
