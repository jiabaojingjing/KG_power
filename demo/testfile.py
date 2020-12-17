from py2neo import Graph, Node, Relationship

for col in range(5):
    print("col "+str(col))
    for j in range(8):
        print("j "+str(j))
        if j==3 and col==2:
            break


