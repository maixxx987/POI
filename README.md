**使用POI SAX事件驱动方式解析Excel文件**

在日常工作中，经常遇到需要解析Excel的场景。本文主要讲述Excel2007（.xlsx）及以上版本的解析方式。

Apache POI提供如下解析方式：

![poi feature](/markdown/image/poi-features.png) 

​	从上图可以看到，POI读取Excel2007以上（尾缀为.xlsx）的文件有两种方式，一种是SAX，一种是DOM。SAX（simple API for XML）是一种XML解析的替代方法。不同于DOM解析XML文档时把所有内容一次性加载到内存中的方式，，SAX是一种速度更快，更有效的方法。它逐行扫描文档，一边扫描一边解析。而且相比于DOM，SAX可以在解析文档的任意时刻停止解析，但任何事物都有其相反的一面，对于SAX来说就是操作复杂。



下面我们来看一个例子，比如我们要解析以下表格：

![sheet1](/markdown/image/sheet1.png)

​	从上图可以看到，有部分单元格内容为空，如果用poi默认的解析方式，将会跳过这些单元格。至于为什么会跳过，看了这份sheet的xml文件即可明白。



首先我们看第3行（行号为3，学号为1）解析为XML后的内容：

![sheet1 row1](/markdown/image/sheet1_row3_xml.png)

如上图所示，我们可以从中得到以下信息：

1. "row"元素代表一行
2. "c"元素代表一个单元格
   - "c"元素的"r"属性的值代表当前的坐标，如"A3"代表第3行A列
   - "c"元素的"t"属性的值代表当前单元格类型，如"s"代表单元格类型为string
3. "v"元素代表单元格中的值



有了以上了解后，我们再来看看第4行（行号为4，学号为2）解析为XML后的内容：

![sheet1 row1](/markdown/image/sheet1_row4_xml.png)

从上图可以看到，空单元格是**直接跳过没有解析**的，如上图所示，少了D,E这两列的内容。

所以我们在处理的时候，要注意这种情况。遇到这种情况，可以通过计算两列之间的差值（在图中即为C-F的差值），然后逐个补充上null。如果不补上差值，原本D4的内容将会被F4覆盖。

这种情况是当前单元格信息被清除内容，或者从来没有填写过造成的。



还有一种情况，如下所示：

![sheet1 row1](/markdown/image/sheet1_row1_xml.png)

这种情况是例子中的第一行，即合并单元格的那行，我们可以看到，B-F是有C元素，但是没有V元素。在这种情况下，POI官方提供的DEMO也是直接跳过的，因为POI是遇到V元素才开始解析值的内容。

这种情况下，我们可以通过记录上一个元素标签，与当前元素标签进行比较。如果上一个标签为C元素，则表示漏了一行，如果上一个标签为V元素，则表示正常记录。



相关网站：

<a href="https://poi.apache.org/">Apache POI 官网</a>

<a href="http://poi.apache.org/apidocs/index.html">Apache POI API</a>

<a href="http://poi.apache.org/components/spreadsheet/how-to.html">Apache POI官网示例</a>

   