from py2neo import Graph, Node, Relationship, NodeMatcher

file_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123"
)

file_graph.delete_all()