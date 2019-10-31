一、文件说明
    ├─  readme.txt # 本文件
    ├─  conf.ini   # 程序的配置文件
    ├─  reckon.exe # 可执行程序文件
    ├─  in         # 程序的默认输入文件夹
    └─  out        # 程序的默认输出文件夹

二、conf.ini
     input_forder=   # 配置输入文件夹的路径，把需要处理的excel文件存放到此文件夹下，只支持.xlsx和xls
     output_forder=  # 配置结果输入文件夹的路径，最终会把结果写到同名的csv文件里，存在会覆盖
     key_col=1       # 输入excel文件里内容的key的列数，从0开始，1表示第2列
     data_col=2      # 货品的列数
     group_count=2   # 组合数

###  notice

# 程序书写比较急躁，python语言也是第一次编写，在本地win7 64bit环境少量的数据下测试通过，关于中文编码等问题没有彻底处理
请尽量避开中文，比如中文路径等，若有问题出现，请联络87893689@qq.com。本exe在使用pyinstaller打包成exe文件时，qq管家可
能会列为木马，虽然代码中没有书写任何恶意代码，请酌情使用。若有损失，本作者不承担任何责任。
pyinstaller下载地址为http://www.pyinstaller.org/downloads.html。版本3.5。