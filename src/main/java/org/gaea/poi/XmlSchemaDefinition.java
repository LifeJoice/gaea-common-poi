package org.gaea.poi;

/**
 * 这是一个对XML SCHEMA中元素名的定义类。
 * <p><b>应该全部为public的静态方法或变量。</b></p>
 * 不要在这里放任何逻辑处理或和XML定义无关的。
 * <p/>为什么是class而不是接口？因为想在这里定义一些静态方法，例如getXXX之类的。可以做些静态转换逻辑（不包含复杂业务的）。
 * Created by Iverson on 2015/7/30.
 */
public class XmlSchemaDefinition {
    public static String ROOT_NODE = "excel-template";
    public static String BLOCK_NAME = "block";
    public static String FIELD_NAME = "field";
}
