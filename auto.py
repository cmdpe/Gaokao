'''
    Copyright 2025 一只酸掉的NaOH

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
'''

import pandas as pd
print(f"免责声明：此程序仅辅助寻找数据，最终是否符合要求请按照山东招生考试院所提供的方法在22号查询结果")
try:
    df = pd.read_excel('submit.xls', usecols=[1, 2, 4], header=None, dtype={0: str, 1: str})
    print(f"成功读取文件，共 {len(df)} 行数据")
except Exception as e:
    print(f"读取文件失败: {e}")
    exit()

try:
    rate = float(input("请输入您的高考位次："))
except ValueError:
    print("输入的不是有效数字，请重新运行程序")
    exit()

index_dict = {}
for idx, row in df.iterrows():
    pro_val = str(row[1]) if pd.notna(row[1]) else ""
    school_val = str(row[2]) if pd.notna(row[2]) else ""
    key = (pro_val[:2], school_val[:4])
    if key not in index_dict:
        index_dict[key] = (row[1], row[2], row[4])
print(f"索引构建完成，共 {len(index_dict)} 个唯一键")
print(f"请按照志愿表从上往下一个一个尝试，直到第一个输出“投档此专业”，此专业为投档专业，退出则在任意一处写“exit”")

while True:
    try:
        school = input("请输入院校代码：").strip()
        pro = input("请输入专业代码：").strip()

        if school.lower() == 'exit' or pro.lower() == 'exit':
            break

        query_key = (pro[:2] ,school[:4])
        print(f"查询键: {query_key}")

        if query_key in index_dict:
            row_data = index_dict[query_key]
            if row_data[2] < rate:
                print("未投档")
            else:
                print("投档此专业")
                print(f"{row_data[1]} {row_data[2]} {row_data[4]}")
        else:
            print("未找到匹配数据")

    except Exception as e:
        print(f"发生错误: {e}. 请重新输入。")

print("程序已退出")