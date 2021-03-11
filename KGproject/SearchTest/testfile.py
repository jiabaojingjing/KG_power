from pyhanlp import *

text = "攻城狮逆袭单身狗，迎娶白富美，走上人生巅峰" # 怎么可能噗哈哈！

print(HanLP.segment(text))

CustomDictionary = JClass("com.hankcs.hanlp.dictionary.CustomDictionary")

CustomDictionary.add("攻城狮") # 动态增加

CustomDictionary.insert("白富美", "nz 1024") # 强行插入

#CustomDictionary.remove("攻城狮"); # 删除词语（注释掉试试）

CustomDictionary.add("单身狗", "nz 1024 n 1")

# 展示该单词词典中的词频统计 展示分词

print(CustomDictionary.get("单身狗"))

print(HanLP.segment(text))

# 增加用户词典,对其他分词器同样有效

# 注意此处,CRF分词器将单身狗分为了n 即使单身狗:"nz 1024 n 1"

CRFnewSegment = HanLP.newSegment("crf")

print(CRFnewSegment.seg(text))




