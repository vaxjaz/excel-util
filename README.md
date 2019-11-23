# excel 动态下载解析工具类
##导出 </br>

  需要导出的实体类字段上加上@ExcelName注解。</br>
  其中对于自定义数字状态以及时间格式化在导出时候需要变为其他格式，增加注解中expression的值，可以是直接可执行代码，也可以调用指定方法进行返回需要的值（方法只能写在该类中。</br>
  对于required属性如果为true会将表头设置为红色。</br>
##导入 </br>

   excel表头中的列名对应类属性ExcelName中的value值，进行动态赋值,如果需要进行同导入相反的逻辑，需要通过deExpression的代码进行反解析。
   
   
### example
   详见代码中user类使用demo