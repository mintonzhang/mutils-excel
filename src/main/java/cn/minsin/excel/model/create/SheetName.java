//package cn.minsin.excel.model.create;
//
//import cn.minsin.excel.enums.ExcelVersion;
//
//import java.lang.annotation.Documented;
//import java.lang.annotation.ElementType;
//import java.lang.annotation.Retention;
//import java.lang.annotation.RetentionPolicy;
//import java.lang.annotation.Target;
//
///**
// * @author: minton.zhang
// * @since: 2020/4/8 22:18
// */
//@Target({ElementType.TYPE})
//@Retention(RetentionPolicy.RUNTIME)
//@Documented
//public @interface SheetName {
//
//    /**
//     * sheet名称
//     *
//     * @return
//     */
//    String value() default "default";
//
//    /**
//     * Excel版本管理
//     *
//     * @return
//     */
//    ExcelVersion version() default ExcelVersion.VERSION_2007;
//
//    /**
//     * 模板路径
//     * @return
//     */
//    String templatePath() default "";
//
//    /**
//     * sheet下标
//     */
//    int sheetIndex() default 0;
//}
