# mongodb_data_to_word
从mongodb中取数据批量输入到word中
环境:
python3.x

from pymongo import MongoClient
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import os
