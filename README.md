# ExcelServiceByGoLang
Excel Writer Service For php or other scripts. 

# Installation
Windows

打开CMD命令行

cd 到当前目录

运行 ./ExcelService.exe -service install 安装服务

./ExcelService.exe -service start 开启服务 也可以在服务管理界面启动

Linux
ssh shell 界面

给予 ExcelServiceForlinux 运行权限  chmod 755 ./ExcelServiceForlinux

给予 Demo/SimpleCreateExcel.php 运行权限 chmod 755 ./Demo/SimpleCreateExcel.php

运行 ./ExcelServiceForlinux -service install 安装服务

./ExcelServiceForlinux -service start 启动服务

# Create Excel

运行 php ./Demo/SimpleCreateExcel.php

# Development purpose

由于php平台的PHPExcel类 在写Excel 大文件时 速度甚慢, 在1万行左右就要占用大量内存及运行时间。
所以将写Excel文件 包装成一个服务 一个可跨平台跨语言的服务。
好处在于
1. 大大提高写Excel文件的效率，降低服务器在写文件时的内存占用。
2. 可跨平台跨语言 任意语言都可以模仿PhpLib中的ExcelWriter写法。
3. 原本想做成一个PHP的扩展类 做成.so文件，但后来发现做扩展局限性太大 首先就是PHP各版本并不兼容
其次只兼容PHP语言. 做成服务只需要开启 socket扩展既可.

# User objects

比如开发大型ERP项目时，需要导出大量数据,phpExcel等脚本化语言性能不足在短时间内导出.

# Restrictions

需要有服务器/空间的操作权限, 因为需要安装服务.
