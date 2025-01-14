package cn.mapleie.tool.util;


import cn.mapleie.tool.annotation.ExcelColumn;
import cn.mapleie.tool.converter.DataConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;

import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.DateUtil;

import java.util.concurrent.atomic.AtomicInteger;

/**
 * @Description Excel导出工具类
 * @Author lipenghui
 * @Date 2024/12/18 16:15
 */
public class ExcelImportUtil {

    // 大文件阈值：超过10000行使用流式处理
    private static final int LARGE_FILE_THRESHOLD = 10000;

    // 批处理大小
    private static final int BATCH_SIZE = 1000;

    /**
     * 通用Excel导入方法，支持大文件优化和并行处理
     *
     * @param inputStream Excel文件输入流
     * @param clazz       目标实体类
     * @param fileName    文件名（用于判断文件类型）
     * @return java.util.List<T>
     * @author lipenghui
     * @date 2024/12/18 17:12
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class<T> clazz, String fileName) throws Exception {
        // 检查输入流和目标实体类是否为空
        if (inputStream == null || clazz == null) {
            throw new IllegalArgumentException("输入流或目标实体类不能为空");
        }
        // 使用try-with-resources确保Workbook在导入完成后被正确关闭
        try (Workbook workbook = determineWorkbook(inputStream, fileName)) {
            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);
            // 检查sheet是否为空
            if (sheet == null) {
                throw new RuntimeException("Excel sheet为空");
            }
            // 解析标题行
            Map<String, Integer> headerMap = parseHeaderRow(sheet.getRow(0));
            // 获取总行数
            int totalRows = sheet.getLastRowNum() + 1;
            // 选择处理策略
            return (totalRows > LARGE_FILE_THRESHOLD) ? processLargeFile(sheet, clazz, headerMap, totalRows) : processSmallFile(sheet, clazz, headerMap);
        }
    }

    /**
     * 确定工作簿类型
     *
     * @param inputStream Excel文件输入流
     * @param fileName    文件名
     * @return org.apache.poi.ss.usermod.Workbook
     * @author lipenghui
     * @date 2024/12/18 17:22
     */
    private static Workbook determineWorkbook(InputStream inputStream, String fileName) throws IOException {
        // 根据文件名后缀判断文件类型
        return fileName.toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(inputStream) : new HSSFWorkbook(inputStream);
    }

    /**
     * 解析标题行
     *
     * @param headerRow Excel表格的标题行
     * @return java.util.Map<java.lang.String, java.lang.Integer>
     * @author lipenghui
     * @date 2024/12/18 17:21
     */
    private static Map<String, Integer> parseHeaderRow(Row headerRow) {
        // 创建标题行与列索引的映射
        Map<String, Integer> headerMap = new HashMap<>();
        if (headerRow != null) {
            // 使用lambda表达式遍历单元格
            headerRow.forEach(cell -> {
                if (cell != null) {
                    headerMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
                }
            });
        }
        return headerMap;
    }

    /**
     * 处理小文件（低于阈值）
     *
     * @param sheet     Excel工作表
     * @param clazz     目标实体类
     * @param headerMap 标题行与列索引的映射
     * @return java.util.List<T>
     * @author lipenghui
     * @date 2024/12/18 17:21
     */
    private static <T> List<T> processSmallFile(Sheet sheet, Class<T> clazz, Map<String, Integer> headerMap) throws Exception {
        // 使用ArrayList存储结果
        List<T> resultList = new ArrayList<>();
        // 获取行迭代器
        Iterator<Row> rowIterator = sheet.rowIterator();
        // 跳过标题行
        rowIterator.next();
        // 遍历数据行
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            T instance = createInstanceFromRow(row, clazz, headerMap);
            resultList.add(instance);
        }
        return resultList;
    }

    /**
     * 处理大文件（超过阈值）
     *
     * @param sheet     Excel工作表
     * @param clazz     目标实体类
     * @param headerMap 标题行与列索引的映射
     * @param totalRows 总行数
     * @return java.util.List<T>
     * @author lipenghui
     * @date 2024/12/18 17:20
     */
    private static <T> List<T> processLargeFile(Sheet sheet, Class<T> clazz, Map<String, Integer> headerMap, int totalRows) {
        // 使用线程安全的List来存储结果
        List<T> resultList = Collections.synchronizedList(new ArrayList<>());
        // 使用AtomicInteger来管理行索引
        AtomicInteger rowIndex = new AtomicInteger(1);
        // 计算总批次数
        int totalBatches = (int) Math.ceil((double) totalRows / BATCH_SIZE);
        // 并行处理批次
        List<Thread> threads = new ArrayList<>();
        for (int batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
            // 获取当前批次的起始行和结束行
            int startRow = rowIndex.getAndAdd(BATCH_SIZE);
            int endRow = Math.min(startRow + BATCH_SIZE, totalRows);
            // 创建并启动处理批次的线程
            threads.add(new Thread(() -> {
                try {
                    List<T> batchList = processBatch(sheet, clazz, headerMap, startRow, endRow);
                    resultList.addAll(batchList);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }));
        }
        // 启动所有线程
        threads.forEach(Thread::start);
        // 等待所有线程完成
        for (Thread thread : threads) {
            try {
                thread.join();
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new RuntimeException("处理大文件时被中断", e);
            }
        }
        return resultList;
    }

    /**
     * 批次处理方法
     *
     * @param sheet     Excel工作表
     * @param clazz     目标实体类
     * @param headerMap 标题行与列索引的映射
     * @param startRow  起始行索引
     * @param endRow    结束行索引
     * @return java.util.List<T>
     * @author lipenghui
     * @date 2024/12/18 17:20
     */
    private static <T> List<T> processBatch(Sheet sheet, Class<T> clazz, Map<String, Integer> headerMap, int startRow, int endRow) throws Exception {
        // 存储批次结果的列表
        List<T> batchList = new ArrayList<>();
        // 遍历批次内的行
        for (int rowIndex = startRow; rowIndex < endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                T instance = createInstanceFromRow(row, clazz, headerMap);
                batchList.add(instance);
            }
        }
        return batchList;
    }

    /**
     * 从行创建对象实例
     *
     * @param row       Excel行对象
     * @param clazz     目标实体类
     * @param headerMap 标题行与列索引的映射
     * @return T
     * @author lipenghui
     * @date 2024/12/18 17:19
     */
    private static <T> T createInstanceFromRow(Row row, Class<T> clazz, Map<String, Integer> headerMap) throws Exception {
        // 使用反射创建实例
        T instance = clazz.getDeclaredConstructor().newInstance();
        // 获取所有声明的字段
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            // 获取ExcelColumn注解
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                // 获取列名
                String columnName = annotation.value();
                // 从标题映射中获取列索引
                Integer columnIndex = headerMap.get(columnName);
                if (columnIndex != null) {
                    // 获取单元格
                    Cell cell = row.getCell(columnIndex);
                    // 获取单元格值并转换为字符串
                    String cellValue = getCellValueAsString(cell);
                    // 检查是否为必填列
                    if (annotation.required() && (cellValue == null || cellValue.trim().isEmpty())) {
                        throw new RuntimeException("必填列 '" + columnName + "' 不能为空");
                    }
                    // 使用转换器处理
                    DataConverter<?> converter = annotation.converter().getDeclaredConstructor().newInstance();
                    Object convertedValue = converter.convert(cellValue);
                    // 设置字段值
                    field.setAccessible(true);
                    field.set(instance, convertedValue);
                }
            }
        }
        return instance;
    }

    /**
     * 获取单元格字符串值
     *
     * @param cell Excel单元格对象
     * @return java.lang.String
     * @author lipenghui
     * @date 2024/12/18 17:18
     */
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return null;
        // 根据单元格类型处理
        return switch (cell.getCellType()) {
            // 直接返回字符串值
            case STRING -> cell.getStringCellValue();
            // 判断是否是日期
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    // 日期类型转换为字符串
                    ? new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue())
                    // 数字类型转换为字符串
                    : formatNumericValue(cell.getNumericCellValue());
            // 布尔类型转换为字符串
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            // 公式类型，尝试获取计算结果
            case FORMULA -> getCellFormulaValue(cell);
            default -> null;
        };
    }

    /**
     * 格式化数字值
     *
     * @param numericValue 数字值
     * @return java.lang.String
     * @author lipenghui
     * @date 2024/12/18 17:17
     */
    private static String formatNumericValue(double numericValue) {
        return numericValue == Math.floor(numericValue) ? String.valueOf((long) numericValue) : String.valueOf(numericValue);
    }

    /**
     * 处理公式单元格的值
     *
     * @param cell 包含公式的单元格
     * @return java.lang.String
     * @author lipenghui
     * @date 2024/12/18 17:17
     */
    private static String getCellFormulaValue(Cell cell) {
        try {
            // 获取计算后的值
            return switch (cell.getCachedFormulaResultType()) {
                case NUMERIC -> formatNumericValue(cell.getNumericCellValue());
                case STRING -> cell.getStringCellValue();
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                default -> cell.getCellFormula();
            };
        } catch (Exception e) {
            // 如果计算失败，返回公式本身
            return cell.getCellFormula();
        }
    }
}

