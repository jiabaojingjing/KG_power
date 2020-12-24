from py2neo import Graph, Node, Relationship, NodeMatcher
import xlrd
import os
import time
result = []
row=0
col=0
nodedic={}
equiparray=[]
file_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123"
)
superfoldernode=foldernode=subfoldernode=super_folder=folder=sub_folder=filename=""
document=firsttitle=secondtitle=thirdtitle=fourthtitle=documentnode=firstnode=secondnode=thirdnode =fournode=""
#workbook = xlwt.Workbook(encoding='utf-8')
#worksheet = workbook.add_sheet('My Sheet')
path=r"C:\Users\86136\Desktop\文档内容整理\文档"
def wipe_line_break(str):
    return str.replace("\n", "").replace(" ", "")

def delete_BS(str):
    return str.replace(" ", "")

def obtain_vaild_value(*params):
    if params[-1] == "":
        return obtain_vaild_value(*params[:-1])
    else:

        return params[-1]


def get_allfile(cwd):
    global row
    global col
    global folderflag
    global subfolder
    get_dir = os.listdir(cwd)
    for i in get_dir:
        print(i)
        sub_dir = os.path.join(cwd,i)
        #print(sub_dir)
        if os.path.isdir(sub_dir):
            subfolder=True
            row += 1
            get_allfile(sub_dir)
        else:
            result.append(i)
    return result


def SaveToNeo4j():
    data = xlrd.open_workbook(r"C:\Users\86136\Desktop\油浸变压器\powerspec.xls")
    table = data.sheet_by_index(0)
    global super_folder,folder,sub_folder,filename,superfoldernode, foldernode,subfoldernode,filenode
    for row in range(table.nrows):
        blank_num= 0
        start_value = table.cell_value(row, 0)
        if "主文件夹" == start_value:
            continue
        print(table.ncols)
        for col in range(table.ncols):
            value = table.cell_value(row, col)
            if type(value) == str:
                value1 = wipe_line_break(value)
                value = delete_BS(value1)
            if value == "":
                blank_num += 1
                continue
            if col == 0:
                super_folder = value
                folder=sub_folder=filename=foldernode=subfoldernode=filenode=""
                superfoldernode = Node("folder", name=super_folder)
                file_graph.create(superfoldernode)
            elif col == 1:
                folder = value
                subfolder=filename=subfoldernode=filenode=""
                foldernode = Node("folder", name=folder)
                file_graph.create(foldernode)
                file_graph.create(Relationship(obtain_vaild_value(superfoldernode), "包含", foldernode))
            elif col == 2:
                sub_folder = value
                filename=filenode=""
                subfoldernode = Node("folder", name=sub_folder)
                file_graph.create(subfoldernode)
                file_graph.create(Relationship(obtain_vaild_value(superfoldernode, foldernode), "包含", subfoldernode))
            elif col == 3:
                filename = os.path.splitext(value)[0]
                filenode = Node("file", name=filename)
                file_graph.create(filenode)
                file_graph.create(Relationship(obtain_vaild_value(superfoldernode, foldernode,subfoldernode), "包含", filenode))
    file_graph.run("CREATE CONSTRAINT ON (c:folder) ASSERT c.name IS UNIQUE");
    file_graph.run("CREATE CONSTRAINT ON (d:file) ASSERT d.name IS UNIQUE");

def getallequipword():
    filedata = xlrd.open_workbook(r'C:\Users\86136\Desktop\文档内容整理\设备及部件名词.xlsx')
    filetable = filedata.sheet_by_index(1)
    for row in range(1, filetable.nrows):
        for col in range(filetable.ncols):
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value1 = wipe_line_break(value)
                value = delete_BS(value1)
                equiparray.append(value)
                print(value)

def saveFilecontentToNeo4j(filepath,filename):

    global document,documentnode,nodedic
    filedata = xlrd.open_workbook(filepath)
    filetable = filedata.sheet_by_index(0)
    matcher = NodeMatcher(file_graph)
    #为字典开辟空间
    for dicid in range(filetable.ncols+1):
        nodedic[dicid] = ""

    document = os.path.splitext(filename)[0]
    document = wipe_line_break(document)
    document = delete_BS(document)
    documentnode = matcher.match("file", name=document).first()
    if (documentnode is None):
        documentnode = Node("file", name=document)
        file_graph.create(documentnode)
    nodedic[0] = documentnode

    equipname = filetable.cell_value(1, 0)
    if equipname != "":
        equipname = wipe_line_break(equipname)
        equipname = delete_BS(equipname)
        print(equipname)
        if equipname in equiparray:
            equipnode = matcher.match("powerentity", name=equipname).first()
            if (equipnode is None):
                equipnode = Node("powerentity", name=equipname)
                file_graph.create(equipnode)
        else:
            equipnode = Node("powerentity", name=equipname)
            file_graph.create(equipnode)
        file_graph.create(Relationship(equipnode, "包含", documentnode))
    return
    for row in range(1,filetable.nrows):
        for col in range(1,filetable.ncols):
            value = filetable.cell_value(row, col)
            if type(value) == str:
                value1 = wipe_line_break(value)
                value = delete_BS(value1)
            if value == "":
                continue

            newnode = Node("powerentity", name=value)
            file_graph.create(newnode)
            prenode=""
            for dicid1 in range(col,filetable.ncols,1):
                nodedic[dicid1] = ""
            for dicid2 in range(col-1,-1,-1):
                if dicid2 in nodedic:
                    if(nodedic[dicid2]!=""):
                        prenode=nodedic[dicid2]
                        break
                else:
                    print(str(row)+"   "+str(col))
                    print(str(dicid2))
            #print(prenode)
            if(prenode !=""):
                file_graph.create(Relationship(prenode, "包含", newnode))
                nodedic[col]=newnode

    print(document)

def getinfor():
    info=[]
    #info.append(file_graph.run("MATCH (n1:powerentity{name:\"油浸变压器\"})-[]->(n2)-[]->(n3) where n3.name=~'.*全面巡视.*' RETURN n3.name"))

    info.append(file_graph.run("MATCH p = (:powerentity{name:\"电流互感器\"})-[]->(:powerentity{name:\"油色谱\"})-[]->(:powerentity{name:\"H2大于150μL/L\"})-[]->(n3:powerentity) where n3.name =~'.*检修.*' RETURN n3.name"))
    print(info)

if __name__ == "__main__":

    #get_all(r'C:\Users\86136\Desktop\电力缺陷\通辽调研资料')
    #SaveToNeo4j()


    # num=0
    # getallequipword()
    filepathset=get_allfile(path)
    for item in filepathset:
        # if "油浸式变压器" in item:
        filepath=path+"\\"+item
        saveFilecontentToNeo4j(filepath,item)
            # num=num+1
            # print(str(num))
     # saveFilecontentToNeo4j(r'C:\Users\86136\Desktop\文档内容整理\文档\国家电网公司变电评价管理规定（试行） 第1分册 油浸式变压器（电抗器）精益化评价细则.xlsx','国家电网公司变电评价管理规定（试行） 第1分册 油浸式变压器（电抗器）精益化评价细则')

