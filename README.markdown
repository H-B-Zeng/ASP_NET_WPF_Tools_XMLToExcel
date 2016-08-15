FileName: WPF_XMLToExcel  
參考來源:
https://support.microsoft.com/en-us/kb/307548

Features 1:Excel To XML  
Directions:Single Excel Export multiple XML

Features 2:XML To Excel  
Directions:multiple XML Export Single Excel

---
version 1.0.2   
此次新增測試檔案及免安裝執行檔  
Test Data XAML:multiple Xaml Export Single Excel  
Test Data XML :multiple XML Export Single Excel  
  
免安裝執行檔  
Free Setup Can Use\Tools_Multi-language_Xaml To Excel.zip  
  
---
version 1.0.1  
Features 2:XML To Excel  
1.測試資料:XmlData.xml, XmlData.xlsx   
2.新增功能:可將xml轉成excel, XmlData.xml =>> XmlData.xlsx  
  轉換條件:因為各種xml 或xaml 格式不同，此轉換器只限於原始Excel資料像 dataTable  
  如下圖的XML(Excel)才能轉換

Excel format

| name         | street           | tel  |   |   |
|--------------|------------------|------|---|---|
| Joe   Tester | Baker   street 5 | 09xx |   |   |
|              |                  |      |   |   |
|              |                  |      |   |   |


XML format   
``` 
<?xml version="1.0" encoding="UTF-8"?>
<addresses xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
	   xsi:noNamespaceSchemaLocation='test.xsd'>
  <addessTable>
    <name>Joe Tester</name>
    <street>Baker street 5</street>
	<tel>09xx</tel>
  </addessTable>
</addresses>
```

---------------------------------------

version 1.0.0  
Features 2:XML To Excel  
1.多個XML:動態將多個xml，Merge Column, 彙整成單一Excel，重複欄位不覆蓋  
2.存檔路徑:自動產生匯出路徑，預設會自動代入，可以自行指定其他位置。  
3.防呆:匯出Excel時，未選檔案會提醒。  
3.防呆:如果檔案格式不對會提示格式錯誤。  
4.防呆:如果給的存錄路徑有錯，也會提示路徑有錯  

