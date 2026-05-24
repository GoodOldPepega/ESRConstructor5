# Основы
В первую очередь необходимо создать объект ESRConstructor, в который следует передать экселевскую книгу HSSFWorkbook
```java
ESRConstructor5 esrConstructor5 = new ESRConstructor5(hssfWorkbook);
```
Далее необходимо создать объект отчёта ESRReport (Примеры его создания рассматриваются далее) и записать его в ESRConstructor5, используя метод writeToExcel. В этот метод передаются следующие аргументы:
1. Объект ESRReport
2. Имя листа EXCEL
3. Необходимость использования полосатой раскраски (каждый блок будет визуально отличаться от другого)
```java
esrConstructor5.writeToExcel(esrReport, "First Report", false);
```
После этого останется просто записать книгу EXCEL стандартными средствами Apache poi
```java
hssfWorkbook.write(fileOutputStream);
fileOutputStream.close();
```

# Примеры создания отчётов (объектов ESRReport) 

## Простой отчёт с одной строкой заголовка и несколькими строками информации
```java
ESRReport esrReport = new ESRReport();

//Создание описательной части
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A title for first report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Создание заголовка
ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
ESRRow esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column3", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
esrReport.setEsrHeaderBlock(esrHeaderBlock);

//Создание блока основного содержимого
ESRBlock esrBlock = new ESRTableBlock();
ESRRow esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 1.1"));
esrRow.addCell(new ESRTableCell("Cell 1.2"));
esrRow.addCell(new ESRTableCell("Cell 1.3"));
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 2.1"));
esrRow.addCell(new ESRTableCell("Cell 2.2"));
esrRow.addCell(new ESRTableCell("Cell 2.3"));
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 3.1"));
esrRow.addCell(new ESRTableCell("Cell 3.2"));
esrRow.addCell(new ESRTableCell("Cell 3.3"));
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
```

## Простой отчёт с использованием однострочных блоков ESRSingleRowBlock
```java
ESRReport esrReport = new ESRReport();

//Создание описательной части
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A title for second report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Создание заголовка
ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
ESRRow esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column3", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
esrReport.setEsrHeaderBlock(esrHeaderBlock);

//Создание блока основного содержимого
ESRSingleRowBlock esrBlock = new ESRSingleRowBlock();
esrBlock.addCell("Cell 1.1");
esrBlock.addCell("Cell 1.2");
esrBlock.addCell("Cell 1.3");
esrReport.addBlock(esrBlock);
esrBlock = new ESRSingleRowBlock();
esrBlock.addCell("Cell 2.1");
esrBlock.addCell("Cell 2.2");
esrBlock.addCell("Cell 2.3");
esrReport.addBlock(esrBlock);
esrBlock = new ESRSingleRowBlock();
esrBlock.addCell("Cell 3.1");
esrBlock.addCell("Cell 3.2");
esrBlock.addCell("Cell 3.3");
esrReport.addBlock(esrBlock);
```

## Отчёт с несколькими строками заголовка и объединениями ячеек как в заголовке, так и в блоке основного содержимого
```java
ESRReport esrReport = new ESRReport();

//Создание описательной части
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A first row of title for third report");
esrTitleBlock.addTitleRow("A second row of title for third report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Создание заголовка
ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
ESRRow esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column1.1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2.1", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column3.1", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column2.1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2.2", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column2.3", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
ESRMerge esrMerge = new ESRMerge();
esrMerge.addMergeGroup(0,0,0,1);
esrHeaderBlock.setMerge(esrMerge);
esrReport.setEsrHeaderBlock(esrHeaderBlock);

//Создание блока основного содержимого
ESRBlock esrBlock = new ESRTableBlock();
ESRRow esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 1.1"));
esrRow.addCell(new ESRTableCell("Cell 1.2"));
esrRow.addCell(new ESRTableCell("Cell 1.3"));
esrBlock.addRow(esrRow);
//Создание объединения ячеек внутри объекта ESRBlock. Значение берётся из первой объединяемой ячейки
esrMerge = new ESRMerge();
esrMerge.addMergeGroup(0, 0, 1, 2);
esrBlock.setMerge(esrMerge);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 2.1"));
esrRow.addCell(new ESRTableCell("Cell 2.2"));
esrRow.addCell(new ESRTableCell("Cell 2.3"));
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 3.1"));
esrRow.addCell(new ESRTableCell("Cell 3.2"));
esrRow.addCell(new ESRTableCell("Cell 3.3"));
esrBlock.addRow(esrRow);
esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 3.1.1"));
esrRow.addCell(new ESRTableCell("Cell 3.2.2"));
esrRow.addCell(new ESRTableCell("Cell 3.3.3"));
esrBlock.addRow(esrRow);
esrMerge = new ESRMerge();
esrMerge.addMergeGroup(0, 1, 0, 0);
esrMerge.addMergeGroup(0, 1, 2, 2);
esrBlock.setMerge(esrMerge);
esrReport.addBlock(esrBlock);
```

## Отчёт с ячейками, содержащими разный тип информации
```java
ESRReport esrReport = new ESRReport();

//Создание описательной части
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A title for fifth report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Создание заголовка
ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
ESRRow esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column3", 4000));
esrHeaderRow.addCell(new ESRHeaderCell("Column4", 4000));
esrHeaderRow.addCell(new ESRHeaderCell("Column5", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
esrReport.setEsrHeaderBlock(esrHeaderBlock);

//Создание блока основного содержимого
ESRBlock esrBlock = new ESRTableBlock();
ESRRow esrRow = new ESRTableRow();
ESRTableCell esrTableCell = new ESRTableCell("Cell 1.1");
//Строковый тип ячейки
esrTableCell.setType(ESRCell.CELL_TYPE_STRING);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(101);
//Числовой тип ячейки
esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(true);
//Булевый тип ячейки
esrTableCell.setType(ESRCell.CELL_TYPE_BOOLEAN);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("SUM(B4:B5)");
//Ячейка с формулой
esrTableCell.setType(ESRCell.CELL_TYPE_FORMULA);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("Blank cell");
//Пустая ячейка
esrTableCell.setType(ESRCell.CELL_TYPE_BLANK);
esrRow.addCell(esrTableCell);
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrTableCell = new ESRTableCell("Cell 2.1");
esrTableCell.setType(ESRCell.CELL_TYPE_STRING);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(102);
esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(false);
esrTableCell.setType(ESRCell.CELL_TYPE_BOOLEAN);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("B4+B5");
esrTableCell.setType(ESRCell.CELL_TYPE_FORMULA);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("Blank cell");
esrTableCell.setType(ESRCell.CELL_TYPE_BLANK);
esrRow.addCell(esrTableCell);
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrTableCell = new ESRTableCell("Cell 3.1");
esrTableCell.setType(ESRCell.CELL_TYPE_STRING);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(103);
esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(true);
esrTableCell.setType(ESRCell.CELL_TYPE_BOOLEAN);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("B4*B5");
esrTableCell.setType(ESRCell.CELL_TYPE_FORMULA);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("Blank cell");
esrTableCell.setType(ESRCell.CELL_TYPE_BLANK);
esrRow.addCell(esrTableCell);
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
```


## Отчёт без описательной части и заголовков
```java
ESRReport esrReport = new ESRReport();

//Создание блока основного содержимого
ESRBlock esrBlock = new ESRTableBlock();
ESRRow esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 1.1", 3000));
esrRow.addCell(new ESRTableCell("Cell 1.2", 4000));
esrRow.addCell(new ESRTableCell("Cell 1.3", 5000));
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 2.1"));
esrRow.addCell(new ESRTableCell("Cell 2.2"));
esrRow.addCell(new ESRTableCell("Cell 2.3"));
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
esrBlock = new ESRTableBlock();
esrRow = new ESRTableRow();
esrRow.addCell(new ESRTableCell("Cell 3.1"));
esrRow.addCell(new ESRTableCell("Cell 3.2"));
esrRow.addCell(new ESRTableCell("Cell 3.3"));
esrBlock.addRow(esrRow);
esrReport.addBlock(esrBlock);
```
