# Basics
First, you need to create an ESRConstructor object and pass the HSSFWorkbook Excel workbook to it
```java
ESRConstructor5 esrConstructor5 = new ESRConstructor5(hssfWorkbook);
```
Next, you need to create an ESRReport object (examples of how to do this are provided below) and write it to ESRConstructor5 using the writeToExcel method. The following arguments are passed to this method:
1. ESRReport object
2. EXCEL sheet name
3. Use or not a striped pattern (each block will look different from the others)
```java
esrConstructor5.writeToExcel(esrReport, "First Report", false);
```
After that, all that remains is to save the Excel file using the standard Apache POI tools
```java
hssfWorkbook.write(fileOutputStream);
fileOutputStream.close();
```

# Examples of report creation (ESRReport object) 

## A simple report with a single header row and several rows of information
```java
ESRReport esrReport = new ESRReport();

//Creating descriptive part
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A title for first report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Creating header
ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
ESRRow esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column3", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
esrReport.setEsrHeaderBlock(esrHeaderBlock);

//Creating the main content block
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

## A simple report using ESRSingleRowBlock single-row blocks
```java
ESRReport esrReport = new ESRReport();

//Creating descriptive part
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A title for second report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Creating header
ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
ESRRow esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column3", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
esrReport.setEsrHeaderBlock(esrHeaderBlock);

//Creating the main content block
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

## A report with multiple header rows and merged cells in both the header and the main content section
```java
ESRReport esrReport = new ESRReport();

//Creating descriptive part
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A first row of title for third report");
esrTitleBlock.addTitleRow("A second row of title for third report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Creating header
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

//Creating the main content block
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

## A report with cells containing different types of information
```java
ESRReport esrReport = new ESRReport();

//Creating descriptive part
ESRTitleBlock esrTitleBlock = new ESRTitleBlock();
esrTitleBlock.addTitleRow("A title for fifth report");
esrReport.setEsrTitleBlock(esrTitleBlock);

//Creating header
ESRHeaderBlock esrHeaderBlock = new ESRHeaderBlock();
ESRRow esrHeaderRow = new ESRHeaderRow();
esrHeaderRow.addCell(new ESRHeaderCell("Column1", 3000));
esrHeaderRow.addCell(new ESRHeaderCell("Column2", 3500));
esrHeaderRow.addCell(new ESRHeaderCell("Column3", 4000));
esrHeaderRow.addCell(new ESRHeaderCell("Column4", 4000));
esrHeaderRow.addCell(new ESRHeaderCell("Column5", 4000));
esrHeaderBlock.addRow(esrHeaderRow);
esrReport.setEsrHeaderBlock(esrHeaderBlock);

//Creating the main content block
ESRBlock esrBlock = new ESRTableBlock();
ESRRow esrRow = new ESRTableRow();
ESRTableCell esrTableCell = new ESRTableCell("Cell 1.1");
//String cell
esrTableCell.setType(ESRCell.CELL_TYPE_STRING);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(101);
//Numeric cell
esrTableCell.setType(ESRCell.CELL_TYPE_NUMERIC);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell(true);
//Boolean cell
esrTableCell.setType(ESRCell.CELL_TYPE_BOOLEAN);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("SUM(B4:B5)");
//Formula cell
esrTableCell.setType(ESRCell.CELL_TYPE_FORMULA);
esrRow.addCell(esrTableCell);
esrTableCell = new ESRTableCell("Blank cell");
//Empty cell
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


## Report without a descriptive section or headings
```java
ESRReport esrReport = new ESRReport();

//Creating the main content block
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
