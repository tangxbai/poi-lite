

### POI-LITE

[![poi-lite](https://img.shields.io/badge/plugin-poi--lite-green?style=flat-square)](https://github.com/tangxbai/mybatis-mappe) [![maven central](https://img.shields.io/badge/maven%20central-v1.3.0-brightgreen?style=flat-square)](https://maven-badges.herokuapp.com/maven-central/org.mybatis/mybatis) ![size](https://img.shields.io/badge/size-196kB-green?style=flat-square) [![license](https://img.shields.io/badge/license-Apache%202-blue?style=flat-square)](http://www.apache.org/licenses/LICENSE-2.0.html)



## 项目简介

使用 `poi` 完成 excel 的导入导出功能，封装极简的Api接口供您使用，函数式编程使得调用更为方便和快捷，支持各种自定义操作和丰富的数据转换功能，简洁化的导出样式美化等，支持基于JavaBean的注解式配置，也支持函数式配置驱动，并提供多种导入导出的姿势供您选择。

*注意：此项目是一款完全开源的项目，您可以在任何适用的场景使用它，商用或者学习都可以，如果您有任何项目上的疑问，可以在issue上提出您问题，我会在第一时间回复您，如果您觉得它对您有些许帮助，希望能留下一个您的星星（★），谢谢。*

------

此项目遵照 [Apache 2.0 License]( http://www.apache.org/licenses/LICENSE-2.0.txt ) 开源许可 

技术讨论QQ群：947460272



## 核心亮点

- **兼容性**：兼容高低版本的excel文件格式读写；
- **数据转换**：支持读写操作的数据格式转换；
- **样式美化**：支持自由样式美化操作和隔行高亮标注；
- **支持多Sheet**：支持多sheet读写操作；
- **隐式转换**：Boolean、Enum、Date日期格式等自动化处理；



## 快速开始

Maven方式（**推荐**）

```xml
<dependency>
	<groupId>com.viiyue.plugins</groupId>
	<artifactId>poi-lite</artifactId>
	<version>[VERSION]</version>
</dependency>
```

如果你没有使用Maven构建工具，那么可以通过以下途径下载相关jar包，并导入到你的编辑器。

[点击跳转下载页面](https://search.maven.org/search?q=g:com.viiyue.plugins%20AND%20a:poi-lite&core=gav)

![如何下载](https://user-gold-cdn.xitu.io/2019/10/16/16dd24a506f37022?w=995&h=126&f=png&s=14645)

如何获取最新版本？[点击这里获取最新版本](https://search.maven.org/search?q=g:com.viiyue.plugins%20AND%20a:poi-lite&core=gav)



## 如何使用

1、基于注解

> 模型Bean的配置

```java
@Excel( 
    headerIndex = 0, // 标题行【默认值】
    startIndex = 1, // 内容开始行【默认值】
    cellHeight = 25, // 导出excel时的行高（单位为：点）【默认值】
    strip = true, // 读写标题时是否去掉左右两侧空白【默认值】
    styleable = DefaultStyleable.class // 导出样式美化接口【默认值】
)
public class Bean implements Serializable {

    @ExcelCell( 
        label = "姓名", // 单元格文本
        ignoreHeader = false, // 是否写入操作时忽略所在单元格
        width = 18, // 单元格宽度，数值为这一列可能出现的最大字符数量，默认为0
        widthAutoSize = false, // 是否宽度自适应内容变化【注意：会有效率损耗，慎用】
        dateformat = "yyyy-MM-dd HH:mm:ss", // 日期格式化【默认值】
        bools = { "Y", "N" }, // 当字段类型为布尔类型时，此属性可以便捷的在两者之间进行转换
        styleable = XXXStyleable.class, // 样式美化器【若无则继承至类级别的styleable】
        writer = XXXWriteConverter.class, // 数据写入转换器【默认无】
        reader = DefaultReadConverter.class // 数据读取转换器【默认值】
    )
    private String name;

    @ExcelCell( label = "邮箱地址", width = 24 )
    private String email;

    // 日期格式化支持 Date、LocalDate、LocalDateTime
    @ExcelCell( label = "日期", width = 24, dateformat = "yyyy-MM-dd HH:mm:ss" )
    private Date datetime;

    // 枚举类继承自 EnumConverter 可自动进行数据转换
    // public enum Gender implements EnumConverter<Gender> {}
    @ExcelCell( label = "性别", width = 10 )
    private Gender gender; 

    // 布尔值如果使用 bools 属性可自动进行属性值转换，第一个为true的替代值，第二个为false的替代值
    // 仅能填写两个替代值
    @ExcelCell( label = "状态值", width = 10, bools = { "正常", "异常" } )
    private Boolean status;
    
}
```

> 调用方式

```java
public class Tester {
    
    public static void main( String [] args ) 
    throws IOException, InvalidFormatException {
        
        // 读取
        ExcelReader reader = ExcelReader.of( Bean.class );
        reader.read( .. ).forEach( ( rowNumber, bean ) -> { // 遍历可以获取行号
            System.out.println( rowNumber + " - " + bean );
        } );
        List<Bean> beans = reader.getDataList(); // 此方式无法获取行号
        
        // 写入
        ExcelWriter writer = ExcelWriter.of( Bean.class );
        writer.addSheet( beans );
        writer.addSheet( "Sheet名字", beans );
        writer.append( [Bean] ); // 可以追加数据到最后一个Sheet下
        writer.writeTo( .. ); // 文件、文件路径，数据流等
    }
    
}
```



3、基于配置

> 配置方式

```java
// 基础配置
ExcelInfo<Map<String,Object>> info = ExcelInfo.ofMap()
    .headerIndex( 0 ) // 标题开始行
    .startIndex( 1 ) // 内容开始行【标题和内容之间可以不是连续的行】
    .reader( null ) // 统一读取转换器【默认：DefaultReadConverter】
    .writer( null ) // 写入转换器
    .styleable( null ); // 样式美化器【默认：DefaultStyleable】

// 最简单的列，未配置部分使用默认值
info.cells( "姓名", "邮箱地址", "日期", "性别", "状态值" );

// 更多自定义列选项
// CellInfo.newMapCell( "姓名" ) // 标签【必填】
//    .dateformat( "yyyy-MM-dd" ) // 日期格式化模板【可选】
//    .bools( "正常", "异常" ) // 布尔替换值【可选】
//    .styleable( null ) // 自定义样式美化器【可选】
//    .writer( null ) // 写入转换器【可选】
//    .reader( null ) // 读取转换器【可选】
//    .width( 25 ) // 单元格宽度【可选】
//    .widthAutoSize( true ) // 是否自适应宽度【可选】
//    .ignoreHeader( true ) // 是否忽略写入标题【可选】【仅忽略标题，不忽略内容单元格】
info.addCell( CellInfo.newMapCell( "姓名" ).width( 16 ) );
info.addCell( CellInfo.newMapCell( "邮箱地址" ).width( 25 ) );
info.addCell( CellInfo.newMapCell( "日期" ).width( 30 ).dateformat( "yyyy-MM-dd" ) );
info.addCell( CellInfo.newMapCell( "性别" ).width( 10 ) );
info.addCell( CellInfo.newMapCell( "状态值" ).width( 10 ).bools( "正常", "异常" ) );
```

> 调用方式

```java
public class Tester {
    
    public static void main( String [] args ) 
    throws IOException, InvalidFormatException {
        
        // 读取
        ExcelReader reader = ExcelReader.of( info );
        reader.read( .. ).forEach( ( rowNumber, map ) -> { // 遍历可以获取行号
            System.out.println( rowNumber + " - " + map );
        } );
        List<Map<String, Object> maps = reader.getDataList(); // 此方式无法获取行号
        
        // 写入
        ExcelWriter writer = ExcelWriter.of( info );
        writer.addSheet( maps );
        writer.addSheet( "Sheet名字", maps );
        writer.append( [Map] ); // 可以追加数据到最后一个Sheet下
        writer.writeTo( .. ); // 文件、文件路径，数据流等
    }
    
}
```



## 自定义样式

> 默认样式渲染

```java
// 注意：自定义颜色值仅支持56种颜色，超过56种颜色会默认使用黑色作为设置颜色。
public class DefaultStyleable<T> implements Styleable<T> {

    // 美化Excel导出标题
	@Override
	public Style beautifyHeader( Workbook wb, String label ) {
        // 创建样式对象，第二个参数为样式名，不可重复重复创建会被缓存，可放心使用。
        // 注意：重复对象会被缓存，但每一个不同的样式都需要创建一个新对象才可以
        // 否则具备相同样式的单元格都会具备该样式。
		return Style.of( wb, "header" )
            .alignment( Alignment.CENTER ) // 文字在单元格的排布方位
            .border( BorderStyle.THIN, "#000000") // 设置边框样式、颜色等
            .font( "Microsoft YaHei", 10, true ) // 设置字体样式、大小、粗体、颜色等
            .bgColor( "#D8D8D8" ); // 背景颜色
	}

    // 美化Excel单元格
	@Override
	public Style beautifyIt( Workbook wb, String label, Object v, T e, Integer num ) {
		if ( num % 2 == 0 ) { // 隔行高亮
			return Style.of( wb, "cell:row:" + ( num % 2 ) ).bgColor( "#F5F5F5" );
		}
		return Style.of( wb, "cell" ); // 默认单元格样式
	}

    // 美化每一个单元格
	@Override
	public void applyAll( Style style, String label ) {
		style.alignment( Alignment.LEFT_CENTER );
		if ( style.is( "header" ) ) { // 判断条件即为你自己取的样式名
			style.border( BorderStyle.THIN, "#BFBFBF" );
		} else {
			style.border( BorderStyle.HAIR, "#BFBFBF" );
		}
	}

}
```



## 写入转换器

```java
public class DefaultWriteConverter implements WriteConverter<Bean> {

    public Object convert( Object value ) { // 值不会为空，可直接使用
        if ( value instanceof Long ) {
            return value.toString();
        }
        return value;
    }
    
}
```



## 读取转换器

```java
public class MyReadConverter extends DefaultReadConverter implements ReadConverter {

    public Object readIt( CellInfo<?> info, Class<?> ft, CellType ct, Cell cell ) {
        Object value = super.readIt( info, ft, ct, cell );
        if ( value != null && isType( fieldType, Number.class ) ) {
            return value.toString();
        }
        return value;
    }
    
}
```



## 关于作者

- 邮箱：tangxbai@hotmail.com
- 掘金： https://juejin.im/user/5da5621ce51d4524f007f35f
- 简书： https://www.jianshu.com/u/e62f4302c51f
- Issuse：https://github.com/tangxbai/poi-lite/issues
