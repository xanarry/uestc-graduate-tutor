# uestc-graduate-tutor
电子科技大学研究生导师信息抓取爬虫脚本, 赠送计算机学院导师信息

程序会为每个导师生成一个word文档放在一个文件夹下, 同时会将所有导师的信息放在同一个文本文件中, 方便大家根据关键词快速搜索感兴趣的导师, 避免网页上一个一个打开提高搜索效率

一些说明:
```
0. 作此脚本的目的在于自动生成所有导师的信息汇总, 避免在浏览器中打开海量tab, 不便于查找搜索信息, 爬取地址为http://222.197.183.99/TutorList.aspx
1. 该脚本使用了Python-docx_0.8.6, BeautifulSoup库，使用Python3版本, 执行脚本时请先安装相关库: "pip install python-docx", "pip install bs4" 由于在处理图片时, 该库的原始代码有时出错, 因此请修改:
/usr/local/lib/python3.5/dist-packages/python_docx-0.8.6-py3.5.egg/docx/image/helpers.py
将代码:
unicode_str = chars.decode('UTF-8')
修改为:
unicode_str = chars.decode('UTF-8', errors="ignore")
main.py才不会因此而中断抓取

2. 脚本创建generatefiles目录, 并把对每个导师生成的docx文档和一个txt的文件汇总文件放在该目录中
3. 因为使用代码自动合并所有docx文档到一个文档, 格式会有偏差, 因此没有通过代码生成汇总docx文档, wps可以通过以下方法无损合并文档
    3.1. 打开wps并新建一个空白文档, 然后选择插入, 选择对象, 然后点从文件插入文本, 然后选择全部导师的docx文档即可合成

4. 脚本可以下载所有学院的导师信息, 执行时选择不同的学院编号即可, 此包仅仅包包计算机学院的
5. 目录中包含4个文件, main.py, default.gif, sample_n.docx, sample_w.docx, 分别为脚本, 下载头像失败使用的默认头像, 创建docx文档使用的模板文档, 脚本可修改(会python的话), 模板文件勿动, 否则会出错

```
