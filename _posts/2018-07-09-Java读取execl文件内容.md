---
title:  "Java读取execl文件内容"
toc: true
search: false
permalink: /tech/java-execl
tag: 技术
excerpt: "时间统计第一步"
last_modified_at: 2018-07-08T08:06:00-05:00
typora-copy-images-to: ..\assets\images
---

使用Idea编辑器操作execl文件，首先要引入apache的一个jar包，叫做poi。

# 涉及到知识点

| 知识点                                                       | 疑问                                         |
| ------------------------------------------------------------ | -------------------------------------------- |
| idea如何引入jar包， 打开 File -> Project Structure，单击 Modules -> Dependencies -> "+" -> "Jars or directories"， 选择硬盘上的jar包  ，Apply -> OK | 这样只是在这一个项目中导入，之后可以直接使用 |
|                                                              |                                              |
|                                                              |                                              |

~~~java
try {
            FileInputStream file = new FileInputStream(filePath);
            XSSFWorkbook sheets = new XSSFWorkbook(file);
            int numberOfSheets = sheets.getNumberOfSheets();
            System.out.println(numberOfSheets);
            // XSSFSheet sheet=sheets.getSheetAt(0);
            for (int i = 0; i < numberOfSheets; i++) {
                XSSFSheet sheet = sheets.getSheetAt(numberOfSheets);
                System.out.println("读了第"+i+"张sheet");

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
~~~

| 问题                                   | 原因                            | 解决方案   |
| -------------------------------------- | ------------------------------- | ---------- |
| Sheet index (2) is out of range (0..1) | for循环里没有用i来迭代，而是num | 将num改成i |
|                                        |                                 |            |
|                                        |                                 |            |

