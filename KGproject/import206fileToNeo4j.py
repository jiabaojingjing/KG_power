# import206fileToNeo4j.py
from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os
import math
import re
import openpyxl
from globalFunction import connectNeo4j,wipe_line_break,extractRelation,getpartdic,get_allfile
partdic={}
reldic={}
entitydic={}
equiparray=[]
# def getpartdic():
def saveFilecontentToNeo4j(filepath):

    global document,documentnode,entitydic,reldic,relnum,entitynum,valueType,partdic,filename,provalueType,equiparray,filedata,equipname,filename
    reldic.clear()
    entitydic.clear()
    filedata = xlrd.open_workbook(filepath)
    filetable = filedata.sheet_by_index(0)
    matcher = NodeMatcher(file_graph)
    relnum = 0
    entitynum = 0
    filename=os.path.basename(filepath).replace(" ","").replace(".xlsx","")
    if "检修" in filename:
        fileType="service"
    if "评价" in filename or "策略" in filename:
        fileType="evaluate"
    if "验收" in filename:
        fileType="acceptance"
    if "运维" in filename:
        fileType="operations"

    equipname=filetable.cell_value(1, 0)

    for key in partdic.keys():
        if key not in equiparray:
            equiparray.append(key)

    for row in range(1,filetable.nrows):
        for col in range(0,filetable.ncols):
            # print(str(row),"  ",str(col))
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue

            attribute =  filetable.cell_value(0, col)
            attribute = wipe_line_break(attribute)
            if attribute=="实体":
                name = ""
                desc = ""
                relid = int(col / 2)
                entityid = math.ceil(col / 2)
                #判断实体类型
                valueType = ""
                if value in partdic[equipname]:
                    valueType = "part"
                if value in equiparray:
                    valueType = "equip"


                if len(valueType)==0:
                    valueType = fileType

                if len(reldic)>0 and len(entitydic)>0:
                    if col>0:
                        fathernode=entitydic[entityid-1]
                        relation=reldic[relid-1]
                        if relation!="" and fathernode!="":

                            if valueType=="detail":
                                name="详情"
                                desc=value
                            else:
                                name=value
                        newnode = matcher.match(valueType, name=value, desc=desc,equip=equipname).first()
                        if newnode is None:
                            newnode = Node(valueType, name=name, desc=desc,equip=equipname)
                            file_graph.create(newnode)
                        file_graph.create(Relationship(fathernode, relation, newnode))
                        entitydic[entityid]=newnode
                    elif col==0:
                        reldic.clear()
                        entitydic.clear()
                        newnode = matcher.match("equip", name=value).first()
                        if (newnode is None):
                            newnode = Node("equip", name=value)
                            file_graph.create(newnode)
                        entitydic[entityid] = newnode
                elif len(reldic)==0 and len(entitydic)==0:
                    newnode = matcher.match("equip", name=value).first()
                    if (newnode is None):
                        newnode = Node("equip", name=value)
                        file_graph.create(newnode)
                    entitydic[entityid]=newnode
            elif attribute=="关系":
                reldic[int(col/2)]=value
    savefileKnowledgePoint(filepath)

def savefileKnowledgePoint(filepath):

    filetable = filedata.sheet_by_index(1)
    matcher = NodeMatcher(file_graph)

    if "检修" in filename:
        provalueType = "service"
    if "评价" in filename:
        provalueType = "evaluate"
    if "验收" in filename:
        provalueType = "acceptance"
    if "运维" in filename:
        provalueType = "operations"

    for key in partdic.keys():
        if key not in equiparray:
            equiparray.append(key)

    for row in range(1, filetable.nrows):
        for col in range(0, filetable.ncols):
            # print(str(row),"  ",str(col))
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue

            attribute = filetable.cell_value(0, col)
            attribute = wipe_line_break(attribute)
            if attribute == "实体":
                name = ""
                desc = ""
                relid = int(col / 2)
                entityid = math.ceil(col / 2)
                # 判断实体类型
                valueType = ""
                if value in partdic[equipname]:
                    valueType = "part"
                if value in equiparray:
                    valueType = "equip"
                if (len(valueType) == 0 and len(value) < 20) or provalueType == "evaluate":
                    valueType = provalueType

                if len(valueType) == 0:
                    valueType = "detail"

                if len(reldic) > 0 and len(entitydic) > 0:
                    if col > 0:
                        fathernode = entitydic[entityid - 1]
                        relation = reldic[relid - 1]
                        if relation != "" and fathernode != "":
                            if relation=="起草人":
                                valueType="drafter"
                            elif relation=="起草单位":
                                valueType = "department"
                            if valueType == "detail":
                                name = "详情"
                                desc = value
                            else:
                                name = value
                        newnode = matcher.match(valueType, name=value, desc=desc, equip=equipname).first()
                        if newnode is None:
                            newnode = Node(valueType, name=name, desc=desc, equip=equipname)
                            file_graph.create(newnode)
                        file_graph.create(Relationship(fathernode, relation, newnode))
                        entitydic[entityid] = newnode
                    elif col == 0:
                        reldic.clear()
                        entitydic.clear()
                        newnode = matcher.match("file", name=filename).first()
                        if (newnode is None):
                            newnode = Node("file", name=filename)
                            file_graph.create(newnode)
                        entitydic[entityid] = newnode
                elif len(reldic) == 0 and len(entitydic) == 0:
                    newnode = matcher.match("file", name=filename).first()
                    if (newnode is None):
                        newnode = Node("file", name=filename)
                        file_graph.create(newnode)
                    entitydic[entityid] = newnode
            elif attribute == "关系":
                reldic[int(col / 2)] = value


if __name__ == "__main__":
    global file_graph
    file_graph = connectNeo4j()

    partdic=getpartdic(r"D:\知识图谱\油浸式变压器\equip.xlsx")

    transformerpath = r"D:\知识图谱\文档"
    transformerarray = []
    # transformerarray.append(transformerpath + "\国家电网公司变电评价管理规定（试行） 第1分册 油浸式变压器（电抗器）精益化评价细则.xlsx")
    transformerarray.append(transformerpath + "\国家电网公司变电检修管理规定（试行） 第1分册 油浸式变压器（电抗器）检修细则.xlsx")
    # transformerarray.append(transformerpath + "\国家电网公司变电验收管理规定（试行） 第1分册  油浸式变压器（电抗器）验收细则.xlsx")
    # transformerarray.append(transformerpath + "\国家电网公司变电运维管理规定（试行） 第1分册  油浸式变压器（电抗器）运维细则.xlsx")
    #
    for path in transformerarray:
        saveFilecontentToNeo4j(path)

