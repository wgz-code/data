import os
from win32com import client as wc

#获取目标目录及其子目录下所有指定类型的文件
def gen_file_list(path,ret,file_type=".docx"):
    filelist = os.listdir(path)
    for filename in filelist:
        de_path = os.path.join(path, filename)
        if os.path.isfile(de_path):
            if de_path.endswith(file_type):  # Specify to find the docx file.
                ret.append(de_path)
        else:
            gen_file_list(de_path,ret,file_type)
    return ret

#将指定目录下的doc文档转换为docx文档，转换完成后将doc文档删除
def doc_to_docx(doc_path):
    word = wc.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_path)
    docx_path,filename = os.path.split(doc_path)
    shotname = os.path.splitext(filename)[0]
    docx_path = os.path.join(docx_path,shotname +".docx")
    #txt=4, html=10, docx=16， pdf=17
    doc.SaveAs(docx_path,16)
    doc.Close()
    word.Quit()
    os.remove(doc_path)

#src_path为原始存储doc文档的目录,dest_path为转换成docx后的目录,可自己指定
src_path="C:\luqiuying"
dest_path="C:\software\hh"
if not os.path.exists(dest_path):
      os.makedirs(dest_path)

os.system ("xcopy  /y /s /e %s %s >nul"  %  (src_path, dest_path))
doclist=gen_file_list(dest_path,[],".doc")
print("文档开始转换!")

for file in doclist:
    doc_to_docx(file)

print ("%d个文件转换成功!"%(len(doclist)))
