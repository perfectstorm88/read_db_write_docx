# python小工具：从数据库自动导出表结构到docx(数据库验收文档)

# 需求场景和功能说明
项目验收需要提交《数据库验收文档》, 需要把数据库表结构信息写到word文档中，比较懒，不想手工一个个写；于是顺手用python写了个小工具，源代码参考[read_db_write_docx](https://github.com/perfectstorm88/read_db_write_docx), 实现下列功能：

  - 读取数据库获取表结构信息(支持MySQL、ORACLE、SQLServer)
  - 把表结构信息转换为docx的表格
  - 读取一个docx模板文件，写到word中的指定位置

![image](https://user-images.githubusercontent.com/37857943/95836145-aa074200-0d71-11eb-84d0-ade0168b7ba9.png)


# 运行环境
- python 3.6 及以上
- 依赖库：python-docx、pymysql、pymssql、cx_Oracle

# 使用说明
**下载源码并安装**
```
git clone https://github.com/perfectstorm88/read_db_write_docx
cd read_db_write_docx
pip install -r requirements.txt
```
**修改配置参数：**

- 修改config.yml中的db_info为自己数据库的链接方式
- 修改config.yml中的word_def为自己需要写到word中的表名
- 也可以修改自己的docx模板，修改文档中的锚点(即把表格写到哪个目录下)，对应word_def.anchor参数

**执行程序**

```
python read_db_write_docx.py
```

# 使用样例：

模板样例
![image](https://user-images.githubusercontent.com/37857943/95713613-6b05bd80-0c99-11eb-824d-b9b8c4ad581f.png)


配置参数(config.yml)如下
```yaml
db_info:
  db_type: 'mysql'   #  也可以支持mysql、oracle、sqlserver三种类型
  host: 'localhost' 
  port: 3306 
  user: root
  password: 'root'
  db: 'mytest'
  charset: "utf8"
template: './template/模板：YY系统_数据库结构设计说明书.docx' # 原始docx文件
output: './xx验收文档.docx' # 输出文档
word_def:
  - anchor: 物理结构设计     # 对应原始docx文件中的锚点,即写入到docx文件的哪个位置
    tables:                # 定义了哪些表结构需要导出，如果不定义该字段，则导出所有表
      - dept__医疗机构      # 表名__中文名
      - dw_applyexamine__审核记录表
      - menu__菜单表
      - role__角色表
```
其中`word_def.tables` 定义了哪些表结构需要导出，如果不定义该字段，则默认把该数据库（`db_info.db`）下的所有表结构都导出来

执行`python read_db_write_docx.py`后，输出内容如下:

![image](https://user-images.githubusercontent.com/37857943/95713733-9e484c80-0c99-11eb-9e4f-9f350d61a803.png)



# 参考
- https://python-docx.readthedocs.io/en/latest/user/quickstart.html 
- https://github.com/python-openxml/python-docx/issues/156 在一段后面插入表格
- https://github.com/python-openxml/python-docx/issues/823  如何根据文本内容找到某一段
- https://github.com/python-openxml/python-docx/issues/33  在docx中删除一个段落
- [分享一个MySQL数据库表结构导出word文档最方便的方法](https://www.hotbak.net/key/MYSQL数据库表结构导出成WORD文档.html): 网上的一个小工具，不支持mac，而且只是把表结构导成html，还需要手工调整格式，略麻烦些
- https://stackoverflow.com/questions/33069697/how-to-setup-cell-borders-with-python-docx 增加表边框
