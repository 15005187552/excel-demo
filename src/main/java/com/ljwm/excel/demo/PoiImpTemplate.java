package com.ljwm.excel.demo;

import cn.hutool.core.io.resource.ClassPathResource;
import cn.hutool.core.util.StrUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jdom2.Attribute;
import org.jdom2.DataConversionException;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.input.SAXBuilder;

import java.io.File;
import java.util.List;
import java.util.stream.IntStream;

/**
 * 根据 xml 创建 excel 导入模板
 * Created by yuzhou on 2018/10/21.
 */
@Slf4j
public class PoiImpTemplate {

  public static final float PX = 37F;
  public static final float EM = 267.5F;


  public static void main(String[] args) {

    ClassPathResource resource = new ClassPathResource("student.xml");

    SAXBuilder builder = new SAXBuilder();
    try {
      // 解析导出xml模板
      Document parse = builder.build(resource.getStream());

      Element root = parse.getRootElement();

      log.info("Create import excel template: {}", root.getAttribute("id").getValue());

      String templateName = root.getAttribute("name").getValue();
      // 创建 excel (xls)
      HSSFWorkbook workbook = new HSSFWorkbook();
      // 创建 sheet
      HSSFSheet sheet = workbook.createSheet(templateName);

      // 设置列宽
      Element colGroup = root.getChild("colgroup");
      setColumnWidth(sheet, colGroup);

      int rowNum = 0;
      // 设置标题
      Element title = root.getChild("title");
      rowNum = setTitle(workbook, sheet, title, rowNum);

      // 设置表头
      Element thead = root.getChild("thead");
      rowNum = setTableHead(workbook, sheet, thead, rowNum);

      // 设置数据区域样式
      Element tbody = root.getChild("tbody");
      rowNum = setTableBody(workbook, sheet, tbody, rowNum);

      File file = new File("data/" + root.getAttribute("code").getValue() +".xls");
      if (!file.getParentFile().mkdirs()) {
        log.debug("no directory created");
      }

      workbook.write(file);
      log.info("Excel Template {} is created, total row number: {}", file.getName(), rowNum);

    } catch (Exception e) {
      log.error("Error when create import template", e);
    }

  }

  /**
   * 设置 表格数据区域样式
   * @param workbook
   * @param sheet
   * @param tbody
   * @param rowNum
   * @return
   */
  private static int setTableBody(HSSFWorkbook workbook, HSSFSheet sheet, Element tbody, int rowNum) throws DataConversionException {
    Element tr = tbody.getChild("tr");
    int repeat = tr.getAttribute("repeat").getIntValue();

    List<Element> tds = tr.getChildren("td");
    IntStream.range(0, repeat).forEach((int rowIndex) -> {
      // 创建一行
      HSSFRow row = sheet.createRow(rowNum + rowIndex);

      IntStream.range(0, tds.size()).forEach((int colIndex) -> {
        Element td = tds.get(colIndex);
        // 创建单元格
        HSSFCell cell = row.createCell(colIndex);
        setCellStyle(workbook, sheet, cell, td);
      });
    });

    return rowNum + repeat;
  }

  /**
   * 设置单元格样式
   * @param workbook 工作簿对象
   * @param cell 单元格
   * @param td 单元格设置
   */
  private static void setCellStyle(HSSFWorkbook workbook, HSSFSheet sheet,  HSSFCell cell, Element td) {
    Attribute typeAttr = td.getAttribute("type");
    String type = typeAttr.getValue();

    // 创建单元格数据格式
    HSSFDataFormat format = workbook.createDataFormat();

    // 创建单元格样式
    HSSFCellStyle cellStyle = workbook.createCellStyle();

    if ("NUMERIC".equalsIgnoreCase(type)) { // 数字
      cell.setCellType(CellType.NUMERIC);

      Attribute formatAttr = td.getAttribute("format");
      String formatValue = formatAttr.getValue();
      cellStyle.setDataFormat(format.getFormat(StrUtil.isBlank(formatValue)? "#,##0.00" : formatValue));


    } else if ("STRING".equalsIgnoreCase(type)) { // 字符串
      cell.setCellValue("");

      cell.setCellType(CellType.STRING);

      cellStyle.setDataFormat(format.getFormat("@"));

    } else if ("DATE".equalsIgnoreCase(type)) { // 日期
      cell.setCellType(CellType.NUMERIC);

      cellStyle.setDataFormat(format.getFormat("yyyy-m-d"));
    } else if ("ENUM".equalsIgnoreCase(type)) { // 枚举
      CellRangeAddressList regions = new CellRangeAddressList(cell.getRowIndex(), cell.getRowIndex(),
        cell.getColumnIndex(), cell.getColumnIndex());

      Attribute enumAttr = td.getAttribute("format");
      String enumValue = enumAttr.getValue();

      // 加载下拉列表内容
      DVConstraint constraint = DVConstraint.createExplicitListConstraint(enumValue.split(","));
      // 数据有效性对象
      HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
      sheet.addValidationData(dataValidation);
    }

    cell.setCellStyle(cellStyle);
  }

  /**
   * 设置 表头
   * @param workbook
   * @param sheet
   * @param thead
   */
  private static int setTableHead(HSSFWorkbook workbook, HSSFSheet sheet, Element thead, int rowNum) {
    List<Element> trs = thead.getChildren("tr");
    IntStream.range(0, trs.size()).forEach((int rowIndex) -> {
      Element tr = trs.get(rowIndex);

      // 创建tr所对应的excel行
      HSSFRow row = sheet.createRow(rowNum + rowIndex);
      List<Element> ths = tr.getChildren("th");
      IntStream.range(0, ths.size()).forEach((int colIndex) -> {
        Element th = ths.get(colIndex);
        Attribute valueAttr = th.getAttribute("value");

        // 创建表头所对应的单元格
        HSSFCell cell = row.createCell(colIndex);
        if (valueAttr != null) {
          String value = valueAttr.getValue();
          cell.setCellValue(value); // 设置表头单元格的内容
        }

      });
    });
    return rowNum + trs.size();
  }

  /**
   * 设置标题
   * @param sheet
   * @param title
   */
  private static int setTitle(HSSFWorkbook workbook, HSSFSheet sheet, Element title, int rowNum) {
    List<Element> trs = title.getChildren("tr");
    IntStream.range(0, trs.size()).forEach((int rowIndex) -> {
      Element tr = trs.get(rowIndex);
      // 创建tr所对应的excel行
      HSSFRow row = sheet.createRow(rowNum + rowIndex);

      // 居中样式
      HSSFCellStyle cellStyle = workbook.createCellStyle();
      cellStyle.setAlignment(HorizontalAlignment.CENTER);

      // 字体
      HSSFFont font = workbook.createFont();
      font.setFontName("仿宋_GB2312");
      font.setBold(true);
      font.setFontHeightInPoints((short) 12);
      cellStyle.setFont(font);


      List<Element> tds = tr.getChildren("td");
      IntStream.range(0, tds.size()).forEach((int colIndex) -> {
        Element td = tds.get(colIndex);
        HSSFCell cell = row.createCell(colIndex);

        Attribute rowSapn = td.getAttribute("rowspan");
        Attribute colSapn = td.getAttribute("colspan");
        Attribute value = td.getAttribute("value");
        if (value != null) {
          cell.setCellStyle(cellStyle);
          cell.setCellValue(value.getValue());

          // 合并单元格并居中
          try {
            sheet.addMergedRegion(new CellRangeAddress(
              rowIndex, rowIndex + rowSapn.getIntValue() - 1,
              colIndex, colIndex + colSapn.getIntValue() - 1));
          } catch (Exception e) {
            log.error("Error when add merged region", e);
          }
        }
      });

    });

    return rowNum + trs.size();
  }

  /**
   * 设置列宽
   * @param sheet
   * @param colGroup
   */
  private static void setColumnWidth(HSSFSheet sheet, Element colGroup) {
    List<Element> cols = colGroup.getChildren("col");
    IntStream.range(0, cols.size()).forEach((int i) -> {
      Element col = cols.get(i);
      Attribute width = col.getAttribute("width");
      String unit = width.getValue().replaceAll("[0-9,\\.]", "");
      String value = width.getValue().replaceAll(unit, "");

      int widthValue = 0;
      if (StrUtil.isBlank(unit) || "px".endsWith(unit)) {
        widthValue = Math.round(Float.parseFloat(value) * PX);
      } else if ("em".endsWith(unit)) {
        widthValue = Math.round(Float.parseFloat(value) * EM);
      }

      sheet.setColumnWidth(i, widthValue);
    });
  }
}
