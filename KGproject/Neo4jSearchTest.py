from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os
from copy import deepcopy
import re
keyworddic={}
kwawarry=[]
fuzzykwdic={}
fuzzykwawarry=[]
file_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123"
)

def wipe_line_break(str):
    return str.replace("\n", "").replace(" ", "")

def getKW(filepath):

    global document,documentnode,nodedic,newflag,keyworddic
    document = ""
    newflag=False
    filedata = xlrd.open_workbook(filepath)
    filetable = filedata.sheet_by_index(0)

    for row in range(0,filetable.nrows):
        for col in range(0,filetable.ncols):
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue
            if col==0:
                if document!=value:
                    if len(kwawarry)>0:
                        keyworddic[document]=deepcopy(kwawarry)
                        kwawarry.clear()
                    document=os.path.splitext(value)[0]
                    newflag=True
                else:
                    newflag=False
            if col==1:
                if newflag:
                    newflag = False
                kwawarry.append(value)
    if document!=None:
        keyworddic[document] = kwawarry
        # print(keyworddic)
        for item in keyworddic:
            list=keyworddic[item]
            # print(item)
            # print(list)
def getFuzzyKW(filepath):

    global document,documentnode,nodedic,newflag,fuzzykwdic
    document = ""
    newflag=False
    filedata = xlrd.open_workbook(filepath)
    filetable = filedata.sheet_by_index(0)

    for row in range(0,filetable.nrows):
        for col in range(0,filetable.ncols):
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue
            if col==0:
                if document!=value:
                    if len(fuzzykwawarry)>0:
                        fuzzykwdic[document]=deepcopy(fuzzykwawarry)
                        fuzzykwawarry.clear()
                    document=os.path.splitext(value)[0]
                    newflag=True
                else:
                    newflag=False
            if col==1:
                if newflag:
                    newflag = False
                fuzzykwawarry.append(value)
    if document!=None:
        fuzzykwdic[document] = fuzzykwawarry
        # print(keyworddic)
        for item in fuzzykwdic:
            list=fuzzykwdic[item]
            print(item)
            print(list)

# 问题中包含文档中所列关键词
def KWmatch(filename,question):
    searchedKWlist=[]
    entitylist = []
    if filename in keyworddic:
        kwlist= keyworddic[filename]
        for item in kwlist:
            match = re.search(item, question)
            if match!=None:
                searchedKWlist.append(item)
        if(len(searchedKWlist)>0):
            for entity in searchedKWlist:
                info="-[*]->(:powerentity{name:"+"'"+entity+"'"+"})"
                entitylist.append(info)
            searchinfo="MATCH p =(:file{name:"+"'"+filename+"'"+"})"+"".join(entitylist)+"-[*]->() return p"
            print(searchinfo)
            print(file_graph.run(searchinfo))
        else:
            fuzzykw=""
            if filename in fuzzykwdic:
                kwlist=fuzzykwdic[filename]
                for item in kwlist:
                    match = re.search(item, question)
                    if match != None:
                        fuzzykw=item
                        break;
                if fuzzykw!="":
                    searchinfo = "MATCH p =(:file{name:" + "'" + filename + "'" + "})-[*]->(n:powerentity) where n.name=~'.*"+fuzzykw+".*' return p"
                    print(searchinfo)
                    print(file_graph.run(searchinfo))




if __name__ == "__main__":
    getKW(r'C:\Users\86136\Desktop\文档内容整理\信息提取(1).xlsx')
    getFuzzyKW(r'C:\Users\86136\Desktop\文档内容整理\模糊的实体名词.xlsx')
    KWmatch(r"国家电网公司变电检修管理规定（试行）第10分册干式电抗器检修细则",r"关键工艺质量控制")
