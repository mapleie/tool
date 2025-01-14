package cn.mapleie.tool.xls.converter;

/**
 * @Description 数据转换器接口--用于将Excel单元格中的字符串值转换为目标类型
 * @Author lipenghui
 * @Date 2024/12/18 16:11
 */

public interface DataConverter<T> {
    /**
     * 将字符串值转换为指定类型
     *
     * @param value 原始字符串值
     * @return 转换后的目标类型值
     * @throws RuntimeException 转换失败时抛出异常
     */
    T convert(String value);
}