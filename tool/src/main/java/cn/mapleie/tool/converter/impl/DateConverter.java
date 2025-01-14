package cn.mapleie.tool.converter.impl;

import cn.mapleie.tool.converter.DataConverter;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @Description 日期数据转换器 --- 将字符串日期转换为 Date 类型
 * @Author lipenghui
 * @Date 2024/12/18 16:18
 */
public class DateConverter implements DataConverter<Date> {
    /**
     * 日期格式
     * 默认使用 yyyy-MM-dd 格式
     */
    private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";

    /**
     * 将字符串日期转换为 Date 对象
     *
     * @param value 日期字符串
     * @return 转换后的 Date 对象
     * @throws RuntimeException 日期转换失败时抛出
     */
    @Override
    public Date convert(String value) {
        try {
            // 使用默认日期格式解析
            SimpleDateFormat sdf = new SimpleDateFormat(DEFAULT_DATE_FORMAT);
            return sdf.parse(value);
        } catch (ParseException e) {
            // 转换失败时抛出运行时异常
            throw new RuntimeException("日期转换错误：" + value, e);
        }
    }
}