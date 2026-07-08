#!/usr/bin/env python
# coding: utf-8

# # Data Extraction (ALMA REPORT API)

# In[12]:


import requests
import xml.etree.ElementTree as ET
import pandas as pd


# In[13]:


api_key = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
base_url = 'https://api-ap.hosted.exlibrisgroup.com/almaws/v1/analytics/reports'

params = {
    'apikey': api_key,
    'path': '/shared/Mahidol University 66MU_INST/Data-Alma-use-statistic-book',
    'limit': 1000,
    'col_names': 'true'
}

all_rows = []
column_mapping = {}


# In[14]:


response = requests.get(base_url, params=params)
while True:
    try:
        root = ET.fromstring(response.content)
    except ET.ParseError as e:
        print(f"เกิดข้อผิดพลาดในการอ่าน XML: {e}")
        break

    # ตรวจสอบสถานะการดึงข้อมูล
    is_finished_elem = [e for e in root.iter() if e.tag.endswith('IsFinished')]
    is_finished = False
    if is_finished_elem and is_finished_elem[0].text:
         is_finished = is_finished_elem[0].text.strip().lower() == 'true'

    # ดึง mapping ของชื่อคอลัมน์จาก Schema
    if not column_mapping:
        for e in root.iter():
            if 'name' in e.attrib and e.attrib['name'].startswith('Column'):
                col_id = e.attrib['name']
                col_name = e.attrib.get('{urn:saw-sql}columnHeading')
                if col_name:
                    column_mapping[col_id] = col_name

    # ดึงข้อมูลแถว
    rows = [e for e in root.iter() if e.tag.endswith('Row')]
    for row in rows:
        row_data = {}
        for col in row:
            # ตัด namespace ออกเพื่อเอาแค่ชื่อ tag (เช่น Column0, Column1)
            col_name = col.tag.split('}')[-1] if '}' in col.tag else col.tag
            row_data[col_name] = col.text

        if row_data:
            all_rows.append(row_data)

    # ถ้าดึงข้อมูลเสร็จแล้วให้ออกลูป
    if is_finished:
        print(f"ดึงข้อมูลเสร็จสิ้น ได้ข้อมูลทั้งหมด {len(all_rows)} แถว")
        break

    # ถ้ายังไม่เสร็จ ให้ดึง Token เพื่อเรียกหน้าถัดไป
    token_elem = [e for e in root.iter() if e.tag.endswith('ResumptionToken')]
    if token_elem and token_elem[0].text:
        token = token_elem[0].text
        print(f"กำลังดึงข้อมูลหน้าถัดไป... (ข้อมูลที่ได้แล้ว: {len(all_rows)} แถว)")
        next_params = {
            'apikey': api_key,
            'token': token
        }
        response = requests.get(base_url, params=next_params)
        if response.status_code != 200:
            print(f"เกิดข้อผิดพลาดในการดึงข้อมูลด้วย Token: {response.status_code}")
            break
    else:
        break



# In[15]:


# นำข้อมูลมาแปลงเป็น DataFrame
if all_rows:
    df = pd.DataFrame(all_rows)

    # เปลี่ยนชื่อคอลัมน์ตามที่ดึงมาจาก XML Schema
    if column_mapping:
        df.rename(columns=column_mapping, inplace=True)


# In[16]:


df.info()


# # Data Extraction (ALMA REPORT API) : Sample XML file

# In[2]:


import xml.etree.ElementTree as ET
import pandas as pd


# In[3]:


all_rows = []
column_mapping = {}


# In[4]:


tree = ET.parse('alma_response.xml')
root = tree.getroot()

# ดึง mapping ของชื่อคอลัมน์จาก Schema
if not column_mapping:
    for e in root.iter():
        if 'name' in e.attrib and e.attrib['name'].startswith('Column'):
            col_id = e.attrib['name']
            col_name = e.attrib.get('{urn:saw-sql}columnHeading')
            if col_name:
                column_mapping[col_id] = col_name

# ดึงข้อมูลแถว
rows = [e for e in root.iter() if e.tag.endswith('Row')]
for row in rows:
    row_data = {}
    for col in row:
        # ตัด namespace ออกเพื่อเอาแค่ชื่อ tag (เช่น Column0, Column1)
        col_name = col.tag.split('}')[-1] if '}' in col.tag else col.tag
        row_data[col_name] = col.text

    if row_data:
        all_rows.append(row_data)


# In[5]:


# นำข้อมูลมาแปลงเป็น DataFrame
if all_rows:
    df = pd.DataFrame(all_rows)

    # เปลี่ยนชื่อคอลัมน์ตามที่ดึงมาจาก XML Schema
    if column_mapping:
        df.rename(columns=column_mapping, inplace=True)


# In[6]:


df.info()


# In[7]:


df.to_excel("test.xlsx")


# In[ ]:




