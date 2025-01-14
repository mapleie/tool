package cn.mapleie.tool.xls;

import cn.mapleie.tool.xls.converter.DataConverter;
import cn.mapleie.tool.xls.converter.impl.DefaultConverter;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description Excel列注解 --用于标注实体类属性与Excel列的映射关系
 * @Author lipenghui
 * @Date 2024/12/18 16:11
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {
    /**
     * Excel列的标题名称
     * 用于匹配Excel表头
     *
     * @return 列标题
     */
    String value() default "";

    /**
     * 是否为必填字段
     * 标记该列是否必须有值
     *
     * @return 是否必填
     */
    boolean required() default false;

    /**
     * 自定义数据转换器
     * 用于将Excel单元格值转换为目标类型
     *
     * @return 转换器类型
     */
    Class<? extends DataConverter> converter() default DefaultConverter.class;
}
