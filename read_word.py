import os
import time
from win32com import client
import pyttsx3

# 初始化Text-to-Speech引擎
engine = pyttsx3.init()

# 打开Word文档
doc_path = input("请输入Word文档的完整路径：")  # "E:\雅思\听力词汇.docx"
word = client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(doc_path)

# 逐行读取并朗读
for paragraph in doc.Paragraphs:
    # 获取段落文本
    text = paragraph.Range.Text
    print(text)  # 在控制台打印文本
    # 使用pyttsx3朗读文本
    engine.say(text)
    engine.runAndWait()  # 等待当前句子读完
    time.sleep(50)  # 暂停5秒

# 关闭文档，不保存更改
doc.Close(SaveChanges=False)

# 退出Word应用程序
word.Quit()
