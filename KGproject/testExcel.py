from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd


test_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123"
)
matcher = NodeMatcher(test_graph)

#KLG_EQUIP = "EQUIP"
KLG_EQUIP_TYPE = "powerentity"
KLG_UNIT = "powerentity"
KLG_UNIT_TYPE = "powerentity"
KLG_PART = "powerentity"


KLG_DEFECT = "powerentity"
KLG_DEFECT_DESC = "powerentity"


DEFECT_CLASSIFY_REL = "分类依据"
DEFECT_REL = "缺陷描述"
UNIT_REL = "部件"
UNIT_TYPE_REL = "部件种类"
PART_REL = "部位"
equip_array=["主变压器","断路器","电压互感器","电流互感器"]
def wipe_line_break(str):
    return str.replace("\n", "")


def obtain_vaild_value(*params):
    if params[-1] == "":
        return obtain_vaild_value(*params[:-1])
    else:
        return params[-1]


data = xlrd.open_workbook(
    r"C:\Users\86136\Desktop\油浸变压器\QGDW1906-2013输变电一次设备缺陷分类标准.xlsx")
table = data.sheet_by_index(0)

equip = equip_node = equip_type = equip_type_node = unit = unit_node = unit_type = unit_type_node = part = part_node = ""


defect_desc = defect_dic = set()
defect = defect_node = ""
start_index = False
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
            if value in equip_array:
                equip=value
            elif value=="":
                if equip not in equip_array:
                    continue
            else:
                equip=value
        if equip not in equip_array:
            continue
        if col == 2:
            print("equip:" + equip)
            print("equiptype:" + value)
            equip_type = value
            equip_type_node = unit = unit_node = unit_type = unit_type_node = part = part_node = ""
            equip_node = Node(KLG_EQUIP_TYPE, name=equip_type)
            test_graph.create(equip_node)
            test_graph.run("CREATE CONSTRAINT ON (c:EQUIP_TYPE) ASSERT c.name IS UNIQUE");
        elif col == 3:
            unit = value
            unit_type = unit_type_node = part = part_node = ""
            if unit != "本体":
                unit_node = Node(KLG_UNIT, name=unit)
                test_graph.create(unit_node)
                test_graph.create(Relationship(obtain_vaild_value(equip_node, equip_type_node),UNIT_REL,unit_node))
        elif col == 4:
            unit_type = value
            part = part_node = ""
            if unit_type!= "本体":
                unit_type_node = Node(KLG_UNIT_TYPE, name=unit_type)
                test_graph.create(unit_type_node)
                test_graph.create(Relationship(obtain_vaild_value(equip_node, equip_type_node, unit_node),UNIT_TYPE_REL,unit_type_node ))
        elif col == 5:
            part = value
            if part != "本体":
                part_node = Node(KLG_PART, name=part)
                test_graph.create(part_node)
                test_graph.create(Relationship(obtain_vaild_value(
                        equip_node, equip_type_node, unit_node, unit_type_node),PART_REL,part_node))
        elif col == 6:
            defect = value
            defect_node = Node(KLG_DEFECT, name=defect)
            test_graph.create(defect_node)
            test_graph.create(Relationship(obtain_vaild_value(equip_node, equip_type_node, unit_node, unit_type_node, part_node),DEFECT_REL,defect_node))
        elif col == 7:
            defect_type = value
        elif col == 8:
            node = ""
            if (defect_type, value) not in defect_desc:
                node = Node(KLG_DEFECT_DESC, name=defect_type, level=value)
                test_graph.create(node)
                one_item = (defect_type, value)
                defect_desc.add(one_item)
            else:
                node = matcher.match(
                    KLG_DEFECT_DESC, name=defect_type, level=value).first()
            test_graph.create(Relationship(defect_node, DEFECT_CLASSIFY_REL, node))
            #test_graph.create(Relationship(obtain_vaild_value(equip_node, equip_type_node, unit_node, unit_type_node, part_node), KLG_PROBLEM_DES, node))

    if blank_num >= 9:
        start_index = False
