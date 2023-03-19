<div align=center>
<img alt="LOGO" src=bill-manager/data/icon.svg  height="256" />

# 个人账单管理(Bill Manager)
<br>
</div>

本`README`由`chatgpt`辅助完成

## 项目简介

`bill_manager`是一个用于管理账单的Python项目，它能方便地导入支付宝、建设银行和微信的账单，并对账单进行分类和统计。此外，它还提供一些便捷的账单管理辅助工具(待开发)。项目通用性强，适用于个人使用。


## 文件结构及功能说明：

- Bill Manager
  - `main.py`: 启动应用的主程序
  - `data`: 数据存储路径
  - `utils`: 存放用于账单管理的具体实现代码的文件夹
    - `util.py`: 支持通用的账单处理函数
    - `alipay.py`: 处理支付宝账单
    - `ccb.py`: 处理建设银行账单
    - `wechat.py`: 处理微信账单



## 如何使用
1. 运行`start.py`或者启动打包后的`Bill Manager.exe`
2. 根据窗口提示选择需要输出的`xlsx`表格地址和原始账单位置
3. 运行结束后会打开之前输出的表格文件查看，账目汇总在工作簿-流水表，没有问题即可关闭，结束本次运行。

### 注意事项：

1. 原始账单地址下需要有微信，支付宝，建行三个文件夹



## 贡献者
@noctboat

如有使用或问题，欢迎联系贡献者。
