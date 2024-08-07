xvba 测试

vscode 安装 xvba: `local-smart.excel-live-server`

左下角 创建 项目 
创建

![](./imgs/1.png)


项目

![](./imgs/2.png)

修改

`config.json`


` "excel_file": "index.xlsb",`

->

` "excel_file": "test_vba.xlsm",`

导入:

![](./imgs/3.png)

错误:

![](./imgs/4.png)

解决方法:

https://blog.csdn.net/JavaSEProgrammer/article/details/125785339

文件---选项---信任中心---信任中心设置---宏设置---勾选启用所有宏---勾选信任对VBA工程对象模型的访问---确定