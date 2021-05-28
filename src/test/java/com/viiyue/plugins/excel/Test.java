package com.viiyue.plugins.excel;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {

	public static void main( String [] args ) {
		String text = "[200028]=\"盾兵具有强大的97防御力和生命力，负重\n较多，移动速度较快187。\",\n";
		Pattern pattern = Pattern.compile( "\\[(\\d+)\\]=\"([^\"]*)\"", Pattern.DOTALL );
		Matcher matcher = pattern.matcher( text );
		while ( matcher.find() ) {
			System.out.println( "序号：" + matcher.group( 1 ) );
			System.out.println( "内容：" + matcher.group( 2 ) );
		}
	}

}
