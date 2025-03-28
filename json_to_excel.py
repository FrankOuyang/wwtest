import json
import pandas as pd
from datetime import datetime

def json_to_excel(input_json, output_excel):
    # 读取JSON文件
    with open(input_json, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 创建Excel写入器
    writer = pd.ExcelWriter(output_excel, engine='xlsxwriter')
    
    # 1. 基本信息表
    basic_info = {
        '项目': ['实验目的', '实验性质', '病原体', '阳性对照', '溶剂对照'],
        '内容': [
            data['purpose']['content'],
            data['nature']['content'],
            data['pathogen']['content'],
            data['positive_control']['content'],
            data['solvent_control']['content']
        ]
    }
    pd.DataFrame(basic_info).to_excel(writer, sheet_name='基本信息', index=False)
    
    # 2. 实验系统表
    exp_sys = {
        '项目': ['数量', '物种', '雄性', '雌性'],
        '内容': [
            data['experimental_sys']['quantity'],
            data['experimental_sys']['species'],
            data['experimental_sys']['male'],
            data['experimental_sys']['female']
        ]
    }
    pd.DataFrame(exp_sys).to_excel(writer, sheet_name='实验系统', index=False)
    
    # 3. 实验方法表
    methods_data = []
    for section in data['methods']['sections']:
        methods_data.append({
            '标题': section['title'],
            '步骤': '\n'.join(section['steps']),
            '内容': section['content']
        })
    pd.DataFrame(methods_data).to_excel(writer, sheet_name='实验方法', index=False)
    
    # 4. 实验分组表
    groups_df = pd.DataFrame(data['experimental_groups']['rows'], 
                            columns=data['experimental_groups']['columns'])
    groups_df.to_excel(writer, sheet_name='实验分组', index=False)
    
    # 5. 检测指标表
    indicators_data = []
    for item in data['detection_indicators']['items']:
        indicators_data.append({
            '指标': item['title'],
            '内容': item['content']
        })
    pd.DataFrame(indicators_data).to_excel(writer, sheet_name='检测指标', index=False)
    
    # 保存Excel文件
    writer.close()
    print(f"Excel文件已生成: {output_excel}")

if __name__ == "__main__":
    input_json = 'wwsolution.json'
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel = f'experiment_protocol_{timestamp}.xlsx'
    json_to_excel(input_json, output_excel)