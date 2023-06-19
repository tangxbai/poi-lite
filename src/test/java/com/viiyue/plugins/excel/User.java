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

import java.io.Serializable;
import java.util.Date;
import java.util.StringJoiner;

import org.apache.poi.ss.usermodel.Workbook;

import com.viiyue.plugins.excel.annotation.Excel;
import com.viiyue.plugins.excel.annotation.ExcelCell;
import com.viiyue.plugins.excel.converter.AbstractConverter;
import com.viiyue.plugins.excel.converter.DefaultStyleable;
import com.viiyue.plugins.excel.converter.WriteConverter;
import com.viiyue.plugins.excel.metadata.Style;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
@Excel( styleable = User.MyStyleable.class )
public class User implements Serializable {

	private static final long serialVersionUID = 1L;

	@ExcelCell( label = "姓名", width = 18 )
	private String name;

	@ExcelCell( label = "年龄", width = 10 )
	private Integer age;

	@ExcelCell( label = "邮箱地址", width = 24 )
	private String email;

	@ExcelCell( label = "日期", width = 24, writer = MyWriteConverter.class )
	private Date datetime;

	@ExcelCell( label = "性别", width = 10 )
	private Gender gender = Gender.FEMALE;

	@ExcelCell( label = "状态值", width = 10, bools = { "正常", "异常" } )
	private Boolean status;

	@Override
	public String toString() {
		return new StringJoiner( ",", "[", "]" ).add( "name=" + name ).add( "age=" + age ).add( "email=" + email )
				.add( "datetime=" + datetime ).add( "gender=" + gender ).add( "status=" + status ).toString();
	}

	public static final User newUser( String name, Integer age, String email, Date datetime, Boolean status ) {
		User user = new User();
		user.setName( name );
		user.setAge( age );
		user.setEmail( email );
		user.setDatetime( datetime );
		user.setStatus( status );
		return user;
	}

	public static class MyWriteConverter extends AbstractConverter implements WriteConverter {
		@Override
		public Object convert( Object value ) {
			return formatter.get().format( ( Date ) value );
		}
	}

	public static class MyStyleable extends DefaultStyleable<User> {
		@Override
		public Style beautifyIt( Workbook wb, String label, Object value, User element, Integer num ) {
			if ( element.getAge() < 10 ) {
				return Style.of( wb, "cell:datetime:" + element.getAge() ).bgColor( "#A8D3B9" );
			}
			return super.beautifyIt( wb, label, value, element, num );
		}
	}

}
