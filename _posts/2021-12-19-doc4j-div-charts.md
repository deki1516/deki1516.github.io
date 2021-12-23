---
layout: post
title: "使用doc4j生成任意自定义word图表"
categories: tools
---

### 什么是 doc4j
docx4j 是一个开源 (ASLv2) Java 库，用于创建和操作 Microsoft Open XML（Word docx、Powerpoint pptx 和 Excel xlsx）文件。
它类似于 Microsoft 的 OpenXML SDK，但适用于 Java。 docx4j 使用 JAXB 来创建内存中的对象，
它的功能非常强大，如果文件格式支持它，你就可以使用 docx4j 来生成。但前提是你需要花时间了解 JAXB的用法 和 Open XML 文档的结构。

### 为什么要用 dc4j

word的格式其实可以用xml来表现, docx4j可以用基于xml的方式来方便的操作docx文档。 


### 如何使用doc4j 

使用doc4j 前，我们需要先了解一个chart在word中是如何被定义、如何被引用的。了解了这些，我们才能更好是利用doc4j提供的功能来实现我们想要的效果。

#### 分析word
首先，将我们需要的指定样式（展示方式、配色等）的图表另存到一个单独的word中
![样例图表](/assets/images/blogs/2021-12-20_224423.png)

修改word后缀为.zip格式，使用zip工具打开压缩包，我们可以发现如下内容 
![word内容定义](/assets/images/blogs/2021-12-22_213216.png)

打开 `word/document.xml`, 可以发现图表在word中的对象属性定义
![word内容定义](/assets/images/blogs/2021-12-22_213743.png)

其中关键内容为 c:chart 节点的  r:id属性， 此处的 `rId4` 对应于 `word/_rels/document.xml.rels` 文件中的内容
![word引用定义](/assets/images/blogs/2021-12-22_214155.png)

至此，我们可以看到一个chart在word中是如何被定义和引用的。那么，下一步，我们就将使用doc4j来创建chart的各个对象。


#### 创建chart的各个对象

##### 1、创建chart的excel数据集
图表的数据集是存放在`word\embeddings`目录下的excel文件中的， 我们可以使用多种途径来创建这个文件，像我们项目组经常使用jxls利用excel模板来创建指定格式的excel文档，还是很便利的，创建excel数据集的过程省略，现在假设我们创建了数据集文件。
![数据集](/assets/images/blogs/2021-12-22_224054.png)

##### 2、创建chart对象
``` java

String chartId = "ageLvel";

// 定义图表xml定义名称
Chart chartPart = new Chart(new PartName("/word/charts/chart_" + chartId + ".xml"));
// 解析模板中图表的xml内容来创建图表对象, 对应 word/charts/chart1.xml 的内容
CTChartSpace ctChartSpace = (CTChartSpace) XmlUtils.unmarshalString(chartXmlString, Context.jc, CTChartSpace.class);
chartPart.setJaxbElement(ctChartSpace);
// 添加图表到word 文档中
wordMLPackage.getMainDocumentPart().addTargetPart(chartPart);
```

#### 3、创建embedded xlsx文件并绑定到chart对象上

``` java
		
// 创建rels文件并与excel绑定

RelationshipsPart owningRelationshipPart = new RelationshipsPart();
owningRelationshipPart.setRelationships(new Relationships());
owningRelationshipPart.setPartName(new PartName("/word/charts/_rels/chart_" + chartId + ".xml.rels"));
Relationship relationship = new Relationship();
relationship.setId("rId1");
relationship.setTarget("../embeddings/Microsoft_Excel_Worksheet_" + chartId + ".xlsx");
relationship.setType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
owningRelationshipPart.addRelationship(relationship);


// 创建excel对象
EmbeddedPackagePart excelObj = new EmbeddedPackagePart(new PartName("/word/embeddings/Microsoft_Excel_Worksheet_" + chartId + ".xlsx"));
excelObj.setContentType(new ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
excelObj.setRelationshipType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
// 构建 chart.xml 和 data.xlsx的对应关系
excelObj.setPackage(wordMLPackage);
// 将excel文件byte内容写入
excelObj.setBinaryData(excelByteArray);
excelObj.setRelationships(owningRelationshipPart);
// 绑定chart.xml 和 embedded
chartPart.addTargetPart(excelObj);

```

#### 4、创建chart的 style和color对象

``` java
// 绑定 chart 和 style
EmbeddedPackagePart styleObj = new EmbeddedPackagePart(new PartName("/word/charts/style_" + chartId + ".xml"));
styleObj.setContentType(new ContentType("application/vnd.ms-office.chartstyle+xml"));
styleObj.setRelationshipType("http://schemas.microsoft.com/office/2011/relationships/chartStyle");
styleObj.setPackage(wordMLPackage);
// 直接附加样例word中 word/charts/style1.xml文件内容
styleObj.setBinaryData(new FileInputStream(new File(styleXmlFilePath)));
styleObj.setRelationships(owningRelationshipPart);
// 绑定chart.xml 和 embedded
chartPart.addTargetPart(styleObj);

// 绑定 chart 和 color
EmbeddedPackagePart colorObj = new EmbeddedPackagePart(new PartName("/word/charts/colors_" + chartId + ".xml"));
colorObj.setContentType(new ContentType("application/vnd.ms-office.chartcolorstyle+xml"));
colorObj.setRelationshipType("http://schemas.microsoft.com/office/2011/relationships/chartColorStyle");
colorObj.setPackage(wordMLPackage);
// 直接附加样例word中 word/charts/colors1.xml文件内容
colorObj.setBinaryData(new FileInputStream(new File(colorsXmlFilePath)));
colorObj.setRelationships(owningRelationshipPart);
// 绑定chart.xml 和 embedded
chartPart.addTargetPart(colorObj);
```

至此，我们已经准备好了一个chart所需要的所有对象，那么下一步，就是决定将chart插入到word中的哪个位置，实际上就是在document.xml文档中的哪个节点append进chart的引用定义。

#### 5、将chart插入word中

首先，找到我们制作的word模板中指定的埋点对象（word模板的制作方式不在本文讨论范围内），将埋点对象替换成图表对象
```java
MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
String rId = "ageLvel";
int chartIndex = 1;
String chartTargetName = "charts/chart_" + ageLvel + ".xml";
String chartXml = "            <w:p xmlns:w =\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n" +
        "            <w:r> \n" +
        "                <w:drawing>\n" +
        "                    <wp:inline xmlns:wp =\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">\n" +
        "                        <wp:extent cx=\"5274310\" cy=\"3076575\"/>\n" +
        "                        <wp:effectExtent l=\"0\" t=\"0\" r=\"2540\" b=\"9525\"/>\n" +
        "                        <wp:docPr id=\"" + chartIndex + "\" name=\"" + chartName + "\"/>\n" +
        "                        <a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n" +
        "                            <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\n" +
        "                                <c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"\n" +
        "                                         xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n" +
        "                                         r:id=\"" + rId + "\"/>\n" +
        "                            </a:graphicData>\n" +
        "                        </a:graphic>\n" +
        "                    </wp:inline>\n" +
        "                </w:drawing>\n" +
        "            </w:r>" +
        "        </w:p>";

P pharse = (P) XmlUtils.unmarshalString(chartXml);

if (pharse == null) {
    return;
}
P toReplace = (P) ((R) text.getParent()).getParent();
int index = mainDocumentPart.getContent().indexOf(toReplace);
mainDocumentPart.getContent().add(index, pharse);
mainDocumentPart.getContent().remove(toReplace);
	
```

至此，我们完成了将目标图表插入到word文档中的操作。

