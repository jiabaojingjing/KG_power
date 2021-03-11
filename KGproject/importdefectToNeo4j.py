from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os
import math
from globalFunction import connectNeo4j,wipe_line_break,extractRelation,getpartdic


KLG_EQUIP = "equip"
KLG_EQUIP_TYPE = "equip"
KLG_UNIT = "part"
KLG_UNIT_TYPE = "part"
KLG_PART = "part"
KLG_DEFECT="defect"
KLG_DEFECT_DESC="detail"
KLG_DEFECT_TYPE="defectType"

DEFECT_DESC = "缺陷描述"
DEFECT_TYPE_REL = "缺陷分类"
EQUIP_REL="设备种类"
DEFECT_CLASSIFY_REL = "分类依据"
DEFECT_REL = "缺陷描述"
UNIT_REL = "部件"
UNIT_TYPE_REL = "部件种类"
PART_REL = "部位"
# equip_array=["主变压器","断路器","电压互感器","电流互感器"]
path=r"D:\知识图谱\油浸变压器\QGDW1906-2013输变电一次设备缺陷分类标准.xlsx"
data = xlrd.open_workbook(path)
table = data.sheet_by_index(0)



filename = os.path.basename(path).replace(" ", "").replace(".xlsx", "")

defect = defect_node = ""

def obtain_vaild_value(*params):
    if params[-1] == "":
        return obtain_vaild_value(*params[:-1])
    else:
        return params[-1]

def associatedFile():
    reldic = {}
    entitydic = {}
    Sheet2=data.sheet_by_index(1)
    for row in range(1,Sheet2.nrows):
        for col in range(Sheet2.ncols):
            value = Sheet2.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                continue
            attribute = Sheet2.cell_value(0, col)
            attribute = wipe_line_break(attribute)
            if attribute == "实体":
                relid = int(col / 2)
                entityid = math.ceil(col / 2)

                if len(reldic) > 0 and len(entitydic) > 0:
                    if col > 0:
                        fathernode = entitydic[entityid - 1]
                        relation = reldic[relid - 1]
                        newnode = matcher.match(KLG_DEFECT_TYPE, name=value).first()
                        if newnode is None:
                            newnode = Node(KLG_DEFECT_TYPE, name=value)
                            test_graph.create(newnode)
                        test_graph.create(Relationship(fathernode, relation, newnode))
                        entitydic[entityid] = newnode
                elif len(reldic) == 0 and len(entitydic) == 0:
                    newnode = matcher.match("file", name=filename).first()
                    if (newnode is None):
                        newnode = Node("file", name=filename)
                        test_graph.create(newnode)
                    entitydic[entityid] = newnode
            elif attribute == "关系":
                reldic[int(col / 2)] = value


def importDectectToNeo4j():
    global start_index,equip , equip_node ,equip_type ,equip_type_node, unit , unit_node , unit_type , unit_type_node, part , part_node,defect,defect_node ,defect_classfy_desc_node,defect_classfy_node
    equip = equip_node = equip_type = equip_type_node = unit = unit_node = unit_type = unit_type_node = part = part_node = defect = defect_node = defect_classfy_desc_node = defect_classfy_node = ""
    start_index = False
    teststr = r"定位块"
    debug=False
    for row in range(table.nrows):
        if not start_index:
            if "序号" == table.cell_value(row, 0):
                start_index = True
            continue
        blank_num, defect_type = 0, ""
        for col in range(table.ncols):
            value = table.cell_value(row, col)
            if type(value) == str:
                value = wipe_line_break(value)
            if value == "":
                blank_num += 1
                continue


            if col == 1:

                equip = value
                equip_node = equip_type_node= unit_node = unit_type_node  = part_node =  defect_classfy_desc_node = defect_classfy_node = defect_node=""
                equip_node = Node(KLG_EQUIP, name=equip)
                test_graph.create(equip_node)
                # test_graph.run("CREATE CONSTRAINT ON (c:EQUIP_TYPE) ASSERT c.name IS UNIQUE");
            if col == 2:
                if value !=teststr and debug:
                    continue
                equip_type_node = unit_node = unit_type_node = part_node = defect_classfy_desc_node = defect_classfy_node = defect_node = ""
                equip_type=value
                equip_type_node = Node(KLG_EQUIP_TYPE, name=equip_type)
                test_graph.create(equip_type_node)
                test_graph.create(Relationship(obtain_vaild_value(equip_node), EQUIP_REL, equip_type_node))

            elif col == 3:
                if value != teststr and debug:
                    continue
                unit = value
                unit_node = unit_type_node = part_node = defect_classfy_desc_node = defect_classfy_node = defect_node = ""
                if unit != "本体":
                    unit_node = Node(KLG_UNIT, name=unit)
                    test_graph.create(unit_node)
                    test_graph.create(Relationship(obtain_vaild_value(equip_node ,equip_type_node),UNIT_REL,unit_node))
            elif col == 4:
                if value != teststr and debug:
                    continue
                unit_type = value
                unit_type_node = part_node = defect_classfy_desc_node = defect_classfy_node = defect_node = ""
                if unit_type!= "本体":
                    unit_type_node = Node(KLG_UNIT_TYPE, name=unit_type)
                    test_graph.create(unit_type_node)
                    test_graph.create(Relationship(obtain_vaild_value(equip_node ,equip_type_node, unit_node),UNIT_TYPE_REL,unit_type_node ))
            elif col == 5:
                # print(str(row)+"   "+value)
                if value != teststr and debug:
                    continue

                # print(obtain_vaild_value(equip_node, equip_type_node, unit_node, unit_type_node, part_node))
                # print("finish")
                part = value
                part_node = defect_node = defect_classfy_desc_node = defect_classfy_node =  ""
                if part != "本体":
                    part_node = Node(KLG_PART, name=part)
                    test_graph.create(part_node)
                    test_graph.create(Relationship(obtain_vaild_value(equip_node ,equip_type_node, unit_node, unit_type_node),PART_REL,part_node))
            elif col == 6:

                if value != teststr and debug:
                    continue
                if equip_type==teststr:
                    print(obtain_vaild_value(equip_node ,equip_type_node, unit_node, unit_type_node, part_node))
                    print("finish")

                defect = value
                defect_classfy_desc_node = defect_classfy_node = defect_node = ""
                defect_node = Node(KLG_DEFECT, name=defect)
                test_graph.create(defect_node)
                test_graph.create(Relationship(obtain_vaild_value(equip_node ,equip_type_node, unit_node, unit_type_node, part_node),DEFECT_REL,defect_node))
            elif col == 7:
                if value != teststr and debug:
                    continue
                defect_classfy_desc = value
                defect_classfy_desc_node = defect_classfy_node =""
                defect_classfy_desc_node=matcher.match(KLG_DEFECT_DESC, name=defect_classfy_desc).first()
                if defect_classfy_desc_node is None:
                    defect_classfy_desc_node = Node(KLG_DEFECT_DESC, name=defect_classfy_desc)
                    test_graph.create(defect_classfy_desc_node)
                test_graph.create(Relationship(obtain_vaild_value(equip_node, equip_type_node, unit_node, unit_type_node, part_node,defect_node),DEFECT_CLASSIFY_REL,defect_classfy_desc_node))

            elif col == 8:
                if value != teststr and  debug:
                    continue
                defect_classfy = value
                defect_classfy_node = matcher.match(KLG_DEFECT_TYPE, name=defect_classfy).first()
                if defect_classfy_node is None:
                    defect_classfy_node = Node(KLG_DEFECT_TYPE, name=defect_classfy)
                    test_graph.create(defect_classfy_node)
                test_graph.create(Relationship(
                    obtain_vaild_value(equip_node, equip_type_node, unit_node, unit_type_node, part_node, defect_node,defect_classfy_desc_node),
                    DEFECT_TYPE_REL, defect_classfy_node))


        if blank_num >= 9:
            start_index = False

    associatedFile()
if __name__ == "__main__":
    global test_graph,matcher

    test_graph = connectNeo4j()
    matcher = NodeMatcher(test_graph)
    importDectectToNeo4j()