# import206fileToNeo4j.py
from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os
import math

from globalFunction import connectNeo4j,wipe_line_break,extractRelation,getpartdic,get_allfile
partdic={}
reldic={}
entitydic={}
equiparray=[]
detctionarray=[]

def saveFilecontentToNeo4j(filepath,fileType):

    global document,documentnode,entitydic,reldic,relnum,entitynum,valueType,partdic,filename,provalueType,equiparray,filedata,equipname
    reldic.clear()
    entitydic.clear()
    filedata = xlrd.open_workbook(filepath)
    filetable1 = filedata.sheet_by_index(0)
    filetable2 = filedata.sheet_by_index(1)

    relnum = 0
    entitynum = 0
    filename=os.path.basename(filepath).replace(" ","").replace(".xlsx","")

    equipname=filetable1.cell_value(1, 0)

    for key in partdic.keys():
        if key not in equiparray:
            equiparray.append(key)

    for row in range(1,filetable1.nrows):
        for col in range(0,filetable1.ncols):
            # print(str(row),"  ",str(col))
            value = filetable1.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue

            attribute =  filetable1.cell_value(0, col)
            attribute = wipe_line_break(attribute)
            if attribute=="实体":
                name = ""
                desc = ""
                relid = int(col / 2)
                entityid = math.ceil(col / 2)
                #判断实体类型
                valueType = getEntityType(value)
                if len(valueType)==0:
                    valueType = fileType

                if len(reldic)>0 and len(entitydic)>0:
                    if col>0:
                        fathernode=entitydic[entityid-1]
                        relation=reldic[relid-1]
                        if relation!="" and fathernode!="":
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
                        newnode = matcher.match(valueType, name=value).first()
                        if (newnode is None):
                            newnode = Node(valueType, name=value)
                            file_graph.create(newnode)
                        entitydic[entityid] = newnode
                elif len(reldic)==0 and len(entitydic)==0:
                    newnode = matcher.match(valueType, name=value).first()
                    if (newnode is None):
                        newnode = Node(valueType, name=value)
                        file_graph.create(newnode)
                    entitydic[entityid]=newnode
            elif attribute=="关系":
                reldic[int(col/2)]=value
    savefileKnowledgePoint(filetable2,fileType)

def savefileKnowledgePoint(filetable2,fileType):

    for row in range(1, filetable2.nrows):
        for col in range(0, filetable2.ncols):
            # print(str(row),"  ",str(col))
            value = filetable2.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue
            attribute = filetable2.cell_value(0, col)
            attribute = wipe_line_break(attribute)
            if attribute == "实体":
                name = ""
                desc = ""
                relid = int(col / 2)
                entityid = math.ceil(col / 2)
                # 判断实体类型
                valueType = getEntityType(value)
                if len(valueType) == 0:
                    valueType = fileType

                if len(reldic) > 0 and len(entitydic) > 0:
                    if col > 0:
                        fathernode = entitydic[entityid - 1]
                        relation = reldic[relid - 1]
                        if relation != "" and fathernode != "":
                            if relation=="起草人":
                                valueType="drafter"
                            elif relation=="起草单位":
                                valueType = "department"
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

  # 判断实体类型是否为设备或部件
def getEntityType(entity):
    if entity in partdic[equipname]:
        entityType = "part"
    if entity in equiparray:
        entityType = "equip"
    return entityType


if __name__ == "__main__":
    global file_graph,matcher
    file_graph = connectNeo4j()
    matcher = NodeMatcher(file_graph)
    partdic=getpartdic(r".\文档\equip.xlsx")

    path = r".\文档"

    servicefilearray=[]
    evaluatefilearray=[]
    acceptancefilearray=[]
    operationsfilearray=[]
    detectionfilearray=[]


    detectionfilearray=get_allfile(path + "\变电检测管理规定细则")
    servicefilearray=get_allfile(path + "\变电检修管理规定细则")
    operationsfilearray= get_allfile(path + "\变电运维管理规定细则")
    acceptancefilearray=get_allfile(path + "\变电验收管理规定细则")
    evaluatefilearray=get_allfile(path + "\变电评价管理规定细则")


    for file in detectionfilearray:
        saveFilecontentToNeo4j(file, "检测")
    for file in servicefilearray:
        saveFilecontentToNeo4j(file, "检修")
    for file in operationsfilearray:
        saveFilecontentToNeo4j(file, "运维")
    for file in acceptancefilearray:
        saveFilecontentToNeo4j(file,"验收")
    for file in evaluatefilearray:
        saveFilecontentToNeo4j(file,"评价")


