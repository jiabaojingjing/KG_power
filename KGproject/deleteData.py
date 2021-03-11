from py2neo import Graph, Node, Relationship, NodeMatcher

file_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123"
)

file_graph.run(r"MATCH (n:`operations`) detach delete n")
file_graph.run(r"MATCH (n:`service`) detach delete n")
file_graph.run(r"MATCH (n:`acceptance`) detach delete n")
file_graph.run(r"MATCH (n:`equip`) detach delete n")
file_graph.run(r"MATCH (n:`evaluate`) detach delete n")
file_graph.run(r"MATCH (n:`part`) detach delete n")
file_graph.run(r"MATCH (n:`defect`) detach delete n")
file_graph.run(r"MATCH (n:`defectType`) detach delete n")
file_graph.run(r"MATCH (n:`detail`) detach delete n")