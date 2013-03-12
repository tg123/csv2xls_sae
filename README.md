CSV to XLS 转换
===============
解决CSV 大数（超过15位） 导入 Excel 时精度丢失的问题，到这里转换后的xls文件再也不会有这样问题了。

运行在[SAE](sae.sina.com.cn) 上 [http://csv2xls.sinaapp.com/](http://csv2xls.sinaapp.com/)


本开开发
========

	# 安装依赖
	saecloud install bottle 
	saecloud install pyexcelerator
	
	# 启动测试服务
	dev_server.py
	

使用的第三方库
==============

 * [py-csv2xls](http://py-csv2xls.sourceforge.net/) 引入时做的改动适应SAE
 * [bootstrap-fileupload](http://jasny.github.com/bootstrap/javascript.html#fileupload)	
