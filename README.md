# A BarTender develope demo in C#
![image](https://github.com/user-attachments/assets/4efc2d84-2e06-4188-8975-b0539b5edf45)

这是一个用c#语言编写的BarTender二次开发程序，展示了BarTender软件的部分功能。

本程序使用的版本是：BarTender Designer 2022 R8 版本。



# 二次开发简介
二次开发须引用```UnityEngine.dll``` 和 ```Seagull.BarTender.Print.dll``` 这两个动态链接库文件。

并且软件还要以.NETFramework 2.0 兼容方式运行，*Demo.exe.config 文件配置如下：

	<?xml version="1.0"?>
	<configuration>
	  <startup useLegacyV2RuntimeActivationPolicy="true">
	    
	  <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.0,Profile=Client"/></startup>
	</configuration>

具体的开发细节就自己去看代码吧！！！
