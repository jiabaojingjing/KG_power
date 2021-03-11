from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os

PMS_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123"
)

matcher = NodeMatcher(PMS_graph)
data = xlrd.open_workbook(r"C:\Users\86136\Desktop\油浸变压器\PMS.xlsx")
table = data.sheet_by_index(0)
entity_set = set()
def PMSdefect():
    for row in range(1,table.nrows):
        node1=node2=""
        entity1_value= table.cell_value(row, 0)
        entity2_value = table.cell_value(row, 2)
        if entity1_value not in entity_set:
            node1 = Node("ENTITY", name=entity1_value)
            PMS_graph.create(node1)
            entity_set.add(entity1_value)
        else:
            node1 = matcher.match("ENTITY", name=entity1_value).first()
        if entity2_value not in entity_set:
            node2 = Node("ENTITY", name=entity2_value)
            PMS_graph.create(node2)
            entity_set.add(entity2_value)
        else:
            node2 = matcher.match("ENTITY", name=entity2_value).first()
        PMS_graph.create(Relationship(node1, table.cell_value(row, 1), node2))

    node3 = Node("ENTITY", name="设备")
    PMS_graph.create(node3)
    equip_set= matcher.match("EQUIP_TYPE").all()
    for tem in equip_set:
        PMS_graph.create(Relationship(node3, "包含", tem))


if __name__ == "__main__":
    PMSdefect()

