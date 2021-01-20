from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os
import math

result = []
row=0
col=0
reldic={}
entitydic={}
relarray=[]
entityarray=[]

file_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123"
)
path=r"C:\Users\86136\Desktop\文档内容整理\文档"
def wipe_line_break(str):
    return str.replace("\n", "").replace(" ", "")

def extractRelation(filepath):
    filedata = xlrd.open_workbook(filepath)
    filetable = filedata.sheet_by_index(0)
    for row in range(1, filetable.nrows):
        for col in range(0, filetable.ncols):
            # print(str(row), "  ", str(col))
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue
            attribute = filetable.cell_value(0, col)
            attribute = wipe_line_break(attribute)
            if attribute == "实体":
               if value not in entityarray and len(value)<15:
                   entityarray.append(value)
                   print(value)
            # if attribute == "关系":
            #     if value not in relarray:
            #         relarray.append(value)
            #         print(value)

def saveFilecontentToNeo4j(filepath):

    global document,documentnode,entitydic,reldic,relnum,entitynum
    reldic.clear()
    entitydic.clear()
    filedata = xlrd.open_workbook(filepath)
    filetable = filedata.sheet_by_index(0)
    matcher = NodeMatcher(file_graph)
    relnum = 0
    entitynum = 0
    for row in range(1,filetable.nrows):
        for col in range(0,filetable.ncols):
            print(str(row),"  ",str(col))
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue
            attribute =  filetable.cell_value(0, col)
            attribute = wipe_line_break(attribute)
            if attribute=="实体":
                relid = int(col/2)
                entityid= math.ceil(col/2)
                if len(reldic)>0 and len(entitydic)>0:
                    if col>0:
                        fathernode=entitydic[entityid-1]
                        relation=reldic[relid-1]
                        if relation!="" and fathernode!="":
                            newnode = matcher.match("powerentity", name=value).first()
                            if (newnode is None):
                                newnode = Node("powerentity", name=value)
                            file_graph.create(Relationship(fathernode, relation, newnode))
                            entitydic[entityid]=newnode
                    elif col==0:
                        reldic.clear()
                        entitydic.clear()
                        newnode = matcher.match("powerentity", name=value).first()
                        if (newnode is None):
                            newnode = Node("powerentity", name=value)
                            file_graph.create(newnode)
                        entitydic[entityid] = newnode
                elif len(reldic)==0 and len(entitydic)==0:
                    newnode = matcher.match("powerentity", name=value).first()
                    if (newnode is None):
                        newnode = Node("powerentity", name=value)
                        file_graph.create(newnode)
                    entitydic[entityid]=newnode
            elif attribute=="关系":
                reldic[int(col/2)]=value




if __name__ == "__main__":


    # saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\文档内容整理\文档\油浸式变压器\国家电网公司变电验收管理规定（试行） 第1分册  油浸式变压器（电抗器）验收细则.xlsx", "")
    # saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\文档内容整理\文档\油浸式变压器\国家电网公司变电检修管理规定（试行） 第1分册 油浸式变压器（电抗器）检修细则.xlsx", "")
    # saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\文档内容整理\文档\油浸式变压器\国家电网公司变电评价管理规定（试行） 第1分册 油浸式变压器（电抗器）精益化评价细则.xlsx", "")
    # saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\文档内容整理\文档\油浸式变压器\国家电网公司变电运维管理规定（试行） 第1分册  油浸式变压器（电抗器）运维细则.xlsx", "")
    saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\油浸式变压器\国家电网公司变电验收管理规定（试行） 第1分册  油浸式变压器（电抗器）验收细则.xlsx")
    saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\油浸式变压器\国家电网公司变电检修管理规定（试行） 第1分册 油浸式变压器（电抗器）检修细则.xlsx")
    saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\油浸式变压器\国家电网公司变电评价管理规定（试行） 第1分册 油浸式变压器（电抗器）精益化评价细则.xlsx" )
    saveFilecontentToNeo4j(r"C:\Users\86136\Desktop\油浸式变压器\国家电网公司变电运维管理规定（试行） 第1分册  油浸式变压器（电抗器）运维细则.xlsx" )
