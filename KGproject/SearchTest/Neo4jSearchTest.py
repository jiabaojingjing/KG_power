from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os
from copy import deepcopy
import re
# import jpype
# from jpype import *
from pyhanlp import *
keyworddic={}
kwawarry=[]
fuzzykwdic={}
fuzzykwawarry=[]
segmentlist=[]
equiplist=[]
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
        # for item in keyworddic:
        #     list=keyworddic[item]
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
        # for item in fuzzykwdic:
        #     list=fuzzykwdic[item]
        #     print(item)
        #     print(list)
def getEquipname(filepath):
    global equiplist
    filedata = xlrd.open_workbook(filepath)
    filetable = filedata.sheet_by_index(0)
    for row in range(0, filetable.nrows):
        for col in range(0, filetable.ncols):
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue
            equiplist.append(value)

# 问题中包含文档中所列关键词
def KWmatch(filename,question):
    searchedKWlist=[]
    entitylist = []
    global segmentlistlength,segmentlist
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
        # else:
        #     fuzzykw=""
        #     if filename in fuzzykwdic:
        #         kwlist=fuzzykwdic[filename]
        #         for item in kwlist:
        #             match = re.search(item, question)
        #             if match != None:
        #                 fuzzykw=item
        #                 break;
        #         if fuzzykw!="":
        #             searchinfo = "MATCH p =(:file{name:" + "'" + filename + "'" + "})-[*]->(n:powerentity) where n.name=~'.*"+fuzzykw+".*' return p"
        #             print(searchinfo)
        #             print(file_graph.run(searchinfo))
        else:

            for term in HanLP.segment(question):
                # print(term)
                if str(term.nature)=="n" or str(term.nature)=="nz":
                    segmentlist.append(term.word)

            segmentlistlength= len(segmentlist)
            while segmentlistlength>0:
                # print(str(segmentlistlength))
                # print(fuzzysearch(segmentlistlength,filename))
                if len(fuzzysearch(segmentlistlength, filename)) >0:
                    break
                segmentlistlength = segmentlistlength - 1

def fuzzysearch(length,filename):
    entitylist = []
    fuzzyconditionlist = []
    for i in range(length):
        label = "n" + str(i)
        entity = "-[*]->(" + label + ":powerentity)"
        fuzzycondition = label + ".name=~'.*" + segmentlist[i] + ".*'"
        entitylist.append(entity)
        fuzzyconditionlist.append(fuzzycondition)
    searchinfo = "MATCH p =(:file{name:" + "'" + filename + "'" + "})" + "".join(entitylist) + " where " + "and ".join(fuzzyconditionlist) + " return p"
    print(searchinfo)
    print(file_graph.run(searchinfo).data())
    return file_graph.run(searchinfo).data()



if __name__ == "__main__":
    getKW(r'C:\Users\86136\Desktop\文档内容整理\信息提取(1).xlsx')
    # getFuzzyKW(r'C:\Users\86136\Desktop\文档内容整理\模糊的实体名词.xlsx')
    # KWmatch(r"国家电网公司变电检修管理规定（试行）第10分册干式电抗器检修细则",r"关键工艺质量控制")


    # print(HanLP.segment('你好，欢迎在Python中调用HanLP的API'))
    # for term in HanLP.segment('下雨天地面积水'):
    #     if str(term.nature)=="n":
    #         print(term.word)  # 获取单词与词性


    # document = "油浸式电力变压器（电抗器）有哪些注意事项"
    # for term in HanLP.segment('分接开关有哪些分类'):
    #     print('{}\t{}'.format(term.word, term.nature))  # 获取单词与词性
    # print(HanLP.extractKeyword(document,2))
    # # 自动摘要
    # print(HanLP.extractSummary(document, 3))
    # 依存句法分析
    # print(HanLP.parseDependency("徐先生还具体帮助他确定了把画雄鹰、松鼠和麻雀作为主攻目标。"))
    KWmatch("国家电网公司变电检修管理规定（试行）第1分册油浸式变压器（电抗器）检修细则", "油浸式变压器引线注意事项")