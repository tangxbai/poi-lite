

### POI-LITE

[![poi-lite](https://img.shields.io/badge/plugin-poi--lite-green?style=flat-square)](https://github.com/tangxbai/mybatis-mappe) [![maven central](https://img.shields.io/badge/maven%20central-v1.3.0-brightgreen?style=flat-square)](https://maven-badges.herokuapp.com/maven-central/org.mybatis/mybatis) ![size](https://img.shields.io/badge/size-196kB-green?style=flat-square) [![license](https://img.shields.io/badge/license-Apache%202-blue?style=flat-square)](http://www.apache.org/licenses/LICENSE-2.0.html)



## 项目简介

使用 `poi` 完成 Excel 的导入导出功能，封装极简的Api接口供您使用，函数式编程使得调用更为方便和快捷，支持各种自定义操作和丰富的数据转换功能，简洁化的导出样式美化等，支持基于JavaBean的注解式配置，也支持函数式配置驱动，并提供多种导入导出的姿势供您选择。

*注意：此项目是一款完全开源的项目，您可以在任何适用的场景使用它，商用或者学习都可以，如果您有任何项目上的疑问，可以在issue上提出您问题，我会在第一时间回复您，如果您觉得它对您有些许帮助，希望能留下一个您的星星（★），谢谢。*

------

此项目遵照 [Apache 2.0 License]( http://www.apache.org/licenses/LICENSE-2.0.txt ) 开源许可 

技术讨论QQ群：947460272



## 核心亮点

- **无侵入**：100%兼容mybatis，不与mybatis冲突，只添加功能，对现有程序无任何影响；



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

1、配置 mybatis.xml

这个文件具体如何配置不作过多说明，你可以拉取相关demo查看详细配置，在配置上也没有什么区别，需要注意的是 **typeAliases**（实体别名配置）一定要配置，不然插件可能无法正常工作。

2、配置数据库实体Bean

```java
@Table( prefix = "t_" )
public class Model {

    

}
```

3、使用方式

```java
SqlSessionFactory factory = new MyBatisMapperFactoryBuilder().build( 
```



## 关于作者

- 邮箱：tangxbai@hotmail.com
- 掘金： https://juejin.im/user/5da5621ce51d4524f007f35f
- 简书： https://www.jianshu.com/u/e62f4302c51f
- Issuse：https://github.com/tangxbai/poi-lite/issues
