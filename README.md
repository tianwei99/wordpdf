# wordpdf
word转pdf demo

如果直接加载本地的word文档，是可以直接转换为pdf的，但是如果使用poi-tl生成word转pdf，会因为poi.xwpf.converter.pdf与poi-tl的依赖包版本冲突而报错

errorlog:
java.lang.NoSuchMethodError: org.apache.poi.POIXMLDocumentPart.getPackageRelationship()Lorg/apache/poi/openxml4j/opc/PackageRelationship;
