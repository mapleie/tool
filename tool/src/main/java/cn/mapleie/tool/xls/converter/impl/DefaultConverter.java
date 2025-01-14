package cn.mapleie.tool.xls.converter.impl;

import cn.mapleie.tool.xls.converter.DataConverter;

/**
 * @Description 默认数据转换器  直接返回原始字符串值，不做任何转换
 * @Author lipenghui
 * @Date 2024/12/18 16:15
 */

public class DefaultConverter implements DataConverter<Object> {
    /**
     * 直接返回原始值，不做任何处理
     *
     * @param value 原始字符串值
     * @return 原始值
     */
    @Override
    public Object convert(String value) {
        return value;
    }
}