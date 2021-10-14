# officeCompression

微软 office 2010+ 文档压缩工具，这终究是个失败的工具，NConvert工具对图片的压缩比远没达到预料的比例，先记录留档，如果有好的图片压缩工具会重新调整。

程序的原理很简单，docx/xlsx/pptx 文件本质上都是压缩包，将文件解压缩后，对文档的图片进行压缩处理后重新打包成相对应的文档。

# 快速开始

### 依赖

+ python >= 3.6
+ NConvert

进入项目目录，使用以下命令安装依赖库

```
pip3 install -r requirements.txt
```

访问 https://www.xnview.com/en/nconvert/ 下载符合自己操作系统的 NConvert 程序

### 使用说明

暂时没有开发命令行

如果需要自定义，修改 `main.py` 文件中的 `FILES_PATH`（文档目录） 和 `COMPR_LEVL`（压缩级别） 即可，不自定义也可以使用默认配置，将文档放到程序根目录让程序自己搜索即可

使用 `$ python3 main.py` 即可运行程序
