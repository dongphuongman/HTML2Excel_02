### 使用说明

---

#### 1.功能

- 爬取网页并解析获取指定内容
- 存储Excel

#### 2.类型

- 存在特定legend标签：解析提取内容
- 不存在特定legend标签：跳过

#### 3.Type

- 存在三种Type：分别存放于Type1，Type2，Type3
- 仅存在两种Type：存放于Type1，Type3，Type2被置空

#### 4.用法

##### 4.1Python解释器运行程序

替换程序入口处的 \#paras参数

- base_dir: Excel Demo所在文件夹路径
- excel_demo: Excel文件名称
- sheet_name: sheet名称

##### 4.2 运行exe程序

若需封装程序，联系作者。

#### 5.结果

- Excel保存到当前程序运行路径的Excels文件夹下
- 出错的网页信息保存到当前程序运行路径的error.log文件中
- 出错指的是：
  - 获取html页面出错，可能是网络不稳定或者爬虫被禁止
  - 寻找特定legend标签出错，可能是没有指定的标签信息或者标签信息格式不同于已知
  - 寻找Type出错，可能是没有Type
  - 寻找特定公司信息出错，可能是没有相应的带有指定标签的公司

#### 6.获取帮助

author：yooongchun

wechat：18217235290

