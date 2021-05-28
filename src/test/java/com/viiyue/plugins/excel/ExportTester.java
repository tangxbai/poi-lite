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

import java.io.IOException;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ExportTester {

	public static void main( String [] args ) throws IOException, InvalidFormatException {
		List<User> dataList = Arrays.asList(
				User.newUser( "tangxbai", 10, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 23, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 60, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 9, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 45, "tangxbai@hotmail.com", new Date(), false ),
				null, // Separtor( Blank line )
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				null, // Separtor( Blank line )
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 58, "tangxbai@hotmail.com", new Date(), true )
		);
		List<User> dataList2 = Arrays.asList(
				User.newUser( "tangxbai", 10, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 11, "tangxbai@hotmail.com", new Date(), true ),
				User.newUser( "tangxbai", 23, "tangxbai@hotmail.com", new Date(), true ),
				null, // Separtor( Blank line )
				User.newUser( "tangxbai", 60, "tangxbai@hotmail.com", new Date(), true )
		);
		ExcelWriter.of( User.class ).addSheet( dataList ).addSheet( dataList2 ).writeTo( "e:\\ccc.xlsx" );
	}

}
