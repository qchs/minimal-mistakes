---
title:  "Java读取execl文件内容（未完更新中）"
toc: true
search: false
permalink: /tech/java-time1
tag: 技术
excerpt: "时间统计第一步"
last_modified_at: 2018-07-14T08:06:00-05:00
typora-copy-images-to: ..\assets\images
---

使用Idea编辑器操作execl文件，首先要引入apache的一个jar包，叫做poi。

# 涉及到知识点

| 知识点                                                       | 疑问                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| idea如何引入jar包， 打开 File -> Project Structure，单击 Modules -> Dependencies -> "+" -> "Jars or directories"， 选择硬盘上的jar包  ，Apply -> OK | 这样只是在这一个项目中导入，之后可以直接使用                 |
| 读完一个execl之后，就是读一张sheet，再是根据行列定位读一个cell，Cell cell=sheet.getRow(j).getCell(i, Row.CREATE_NULL_AS_BLANK); | 先定位到第j行，再去数第i列，当然了，都是从0开始的            |
| cell里面的类型判断使用cell.getCellType()                     | 0-数值，1-字符串，2-公式，3-空值，4-布尔，5-错误             |
| 改变值就是重新写入，double v = cell.getNumericCellValue() * 60; cell.setCellValue(v); | 如此set也不会改变execl值                                     |
| 要来读取日期并将日期进行修改，格式统一为类似180701、180720   |                                                              |
| 1.20要在execl文件中存为文本格式，若是小数点两位，使用程序时会成为1.2 |                                                              |
| 1.20如何构造成180120？1、字符串分割，以.为分隔符2、字符串拼接，不能直接使用18，到明年之后扩展性不好 | String的split方法如何用                                      |
| 日期是使用我自己的180714这种格式，还是java自带的日期格式呢？ | 这个要结合数据库查询时，时间的划分，比如用自带的，是否就可以说2018-4-5至2018-7-5？暂定存为标准格式2017-08-22 |
| 日期格式调整完后，在下一次运行到分割时就会报错，要让已经处理过的不再处理第二次 | 有能判断一个字符串中是否有某个特殊字符吗。 java.lang.String.contains() 方法返回true，当且仅当此字符串包含指定的char值序列 |

~~~java
{
        String filePath = "D:\\6.studio\\dependence\\test1.xlsx";
        try {
            FileInputStream file = new FileInputStream(filePath);
            XSSFWorkbook sheets = new XSSFWorkbook(file);
            int numberOfSheets = sheets.getNumberOfSheets();
            System.out.println("sheet总数： " + numberOfSheets);
            XSSFSheet sheet = sheets.getSheetAt(1);
            //sheet的名称能否拿到？
            //共有多少行
            int rowNum = sheet.getLastRowNum() - sheet.getFirstRowNum();
            //共有多少列
            int columnNum = sheet.getRow(0).getLastCellNum() - sheet.getRow(0).getFirstCellNum();
            //只取类别和时间，并将分数表示的小时，转换为分钟
            for (int j = 2; j <= rowNum; j++) {
                for (int i = 0; i < 8; i++) {
                    Cell cell = sheet.getRow(j).getCell(i, Row.CREATE_NULL_AS_BLANK);
                    //多加一个判断，避免多次运行程序时都乘以60
                    if (cell.getCellType() == 0 && cell.getNumericCellValue() < 24) {
                        double v = cell.getNumericCellValue() * 60;
                        cell.setCellValue(Math.round(v));
                    }
                    System.out.print(cell.toString() + "    ");
                }
                System.out.println();
            }
            FileOutputStream fout = new FileOutputStream(filePath);
            sheets.write(fout);
            fout.flush();
            fout.close();
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
//结果
sheet总数： 3
睡觉    390.0    510.0    510.0    510.0    500.0    510.0    540.0    
吃饭    150.0    105.0    100.0    80.0    170.0    90.0    105.0    
洗漱    120.0    75.0    70.0    90.0    65.0    90.0    80.0    
阅读    540.0    480.0    530.0    590.0    370.0    240.0    460.0    
运动    60.0    150.0    80.0    70.0    15.0    30.0    90.0    
娱-单    60.0    90.0    60.0    60.0    40.0    420.0    45.0    
闲-多    0.0    0.0    60.0    40.0    40.0    60.0    30.0   

~~~

| 问题                                                         | 原因                                                       | 解决方案                                                     |
| ------------------------------------------------------------ | ---------------------------------------------------------- | ------------------------------------------------------------ |
| Sheet index (2) is out of range (0..1)                       | for循环里没有用i来迭代，而是num                            | 将num改成i                                                   |
| 8 1/2 *60=500.00000000000006，为什么其他的又是500呢          | 小数计算是不精确的                                         | 强制保留0位小数，四舍五入Math.round(v)-还是有一位小数，想去掉，转换为int类型 |
| set应该是会改变execl文件里的数据，为何再打开还是分数形式？   | 要强制写入（做一个判断才好，如果已经改变过，就不再乘以60） | FileOutputStream fout   = new FileOutputStream(filePath); sheets.write(fout); fout.flush(); fout.close(); file.close(); |
| date1.split(".")，没有分割出两个字符串。在java.lang包中有String.split()方法，返回是一个数组。 | “.”和“\|”都是转义字符，必须得加"\\";                       | 如果用“.”作为分隔的话，必须是如下写法：  String.split("\\."),如果用“\|”作为分隔的话，必须是如下写法：  String.split("\\|"),这样才能正确的分隔开，不能用String.split("\|"); |
|                                                              |                                                            |                                                              |
|                                                              |                                                            |                                                              |

