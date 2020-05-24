#!/usr/bin/env python
# encoding: utf-8
"""
@version: python3.7
@author: JYFelt
@license: Apache Licence 
@contact: JYFelt@163.com
@site: https://blog.csdn.net/weixin_38034182
@software: PyCharm
@file: __init__.py.py
@time: 2020/5/24 16:11
"""

from pymongo import MongoClient
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import os

client = MongoClient('127.0.0.1', 27017)
collection_users = client["gms"]["users"]
collection_ship = client["gms"]["partnership"]
collection_title = client["gms"]["title"]

query = {"time": {"$gte": datetime(2019, 12, 31)}}
ship_data = collection_ship.find(query)

for _ in ship_data:
    doc = DocxTemplate(r"C:/Users/11985/PycharmProjects/数据处理/data/mongdb_gen/assignmentBook.docx")
    title_name = _['title_name']
    stu_name = pd.DataFrame(collection_users.find({'_id': _['_id']}))['name'].values[0]
    class_name = pd.DataFrame(collection_users.find({'_id': _['_id']}))['class_name'].values[0]
    teach_name = pd.DataFrame(collection_users.find({'_id': _['tea_id']}))['name'].values[0]
    user_name = pd.DataFrame(collection_users.find({'_id': _['_id']}))['username'].values[0]
    content = pd.DataFrame(collection_title.find({'_id': _['title_id']}))['summary'].values[0]
    require = pd.DataFrame(collection_title.find({'_id': _['title_id']}))['require'].values[0]

    print(title_name,
          class_name,
          teach_name,
          user_name,
          content,
          require)

    context = {'title': title_name,
               'name': stu_name,
               'username': user_name,
               'class': class_name,
               'content': content,
               'require': require}
    for x in context.values():
        print(x)
    doc.render(context, autoescape=True)

    r = r'C:/Users/11985/PycharmProjects/数据处理/data/mongdb_gen/任务书'
    path = r + '/' + teach_name
    path = path.strip()
    path = path.rstrip("/")
    isExists = os.path.exists(path)

    if not isExists:
        os.makedirs(path)
    doc.save(path + '/' + stu_name + "宝宝的毕设任务书.docx")

print('done!')
